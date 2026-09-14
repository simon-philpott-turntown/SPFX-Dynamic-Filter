/**
 * @file TaxonomyService.ts
 * @description Service for querying SharePoint Global Term Store taxonomy directly in SharePoint Online (SPFx).
 * Exclusively loads live terms from SharePoint Online Term Store v2.1 REST endpoints,
 * traversing Group -> Term Sets -> Terms and child terms with complete synonym resolution.
 * No mock or dummy terms are injected.
 */

import { ITermStoreTag } from '../models/IContainerModels';

export interface ITermGroup {
  id: string;
  name: string;
  termSets: ITermSet[];
}

export interface ITermSet {
  id: string;
  name: string;
  groupId: string;
  groupName: string;
  terms: ITermStoreTag[];
}

export class TaxonomyService {
  /**
   * Loaded live groups from SharePoint Online.
   * Initialised empty: strictly populated from live SharePoint Term Store.
   */
  private static _cachedGroups: ITermGroup[] = [];

  private static _isFetched: boolean = false;
  private static _fetchPromise: Promise<void> | null = null;

  /**
   * Queries the live SharePoint Online Term Store v2.1 REST API.
   * 1. Fetches all Term Groups: `/_api/v2.1/termStore/groups?$expand=sets($expand=terms)`
   * 2. If a group's sets are not expanded inline, queries `/_api/v2.1/termStore/groups/{groupId}/sets`
   * 3. For each set, queries its terms `/_api/v2.1/termStore/sets/{setId}/terms`
   * 4. Recursively resolves child terms and synonyms.
   */
  public static async initializeFromSharePoint(siteUrl?: string, forceRefresh: boolean = false): Promise<void> {
    if (this._isFetched && !forceRefresh) return;
    if (this._fetchPromise && !forceRefresh) return this._fetchPromise;

    this._fetchPromise = (async () => {
      if (typeof window === 'undefined') return;

      try {
        const spCtx = (window as unknown as { _spPageContextInfo?: { webAbsoluteUrl?: string } })._spPageContextInfo;
        const targetUrl = siteUrl || (spCtx && spCtx.webAbsoluteUrl) || '';

        if (!targetUrl) {
          console.warn('[TaxonomyService] No SharePoint site URL available for live Term Store query.');
          return;
        }

        const headers = {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json'
        };

        // Step 1: Query all term groups with sets expanded
        const groupsEndpoint = `${targetUrl}/_api/v2.1/termStore/groups?$expand=sets`;
        const groupsResp = await fetch(groupsEndpoint, {
          headers,
          credentials: 'include'
        });

        if (!groupsResp.ok) {
          console.warn(`[TaxonomyService] Failed to query term groups (${groupsResp.status}):`, await groupsResp.text());
          return;
        }

        const groupsData = await groupsResp.json();
        if (!groupsData || !Array.isArray(groupsData.value)) {
          return;
        }

        const liveGroups: ITermGroup[] = [];

        for (const g of groupsData.value as Array<{
          id: string;
          displayName?: string;
          name?: string;
          sets?: Array<{
            id: string;
            localizedNames?: Array<{ name: string }>;
          }>;
        }>) {
          const groupName = g.displayName || g.name || 'Group';
          let termSets: ITermSet[] = [];

          // Query sets for this group
          let setsList = Array.isArray(g.sets) && g.sets.length > 0 ? g.sets : [];
          if (setsList.length === 0) {
            try {
              const setsEndpoint = `${targetUrl}/_api/v2.1/termStore/groups/${g.id}/sets`;
              const setsResp = await fetch(setsEndpoint, { headers, credentials: 'include' });
              if (setsResp.ok) {
                const setsData = await setsResp.json();
                if (setsData && Array.isArray(setsData.value)) {
                  setsList = setsData.value;
                }
              }
            } catch (setErr) {
              console.warn(`[TaxonomyService] Error fetching sets for group ${groupName}:`, setErr);
            }
          }

          for (const s of setsList) {
            const setName = s.localizedNames && s.localizedNames[0] ? s.localizedNames[0].name : 'Term Set';
            // Always fetch terms for set with expand=relations,children to capture full hierarchy and parent IDs
            const terms = await TaxonomyService._fetchTermsForSet(targetUrl, s.id, setName, groupName, headers);
            termSets.push({
              id: s.id,
              name: setName,
              groupId: g.id,
              groupName: groupName,
              terms: terms
            });
          }

          liveGroups.push({
            id: g.id,
            name: groupName,
            termSets: termSets
          });
        }

        // Exclusively assign live groups returned from the tenant Term Store
        this._cachedGroups = liveGroups;
      } catch (err) {
        console.error('[TaxonomyService] Live Term Store fetch error:', err);
      } finally {
        this._isFetched = true;
        this._fetchPromise = null;
      }
    })();

    return this._fetchPromise;
  }

  /**
   * Helper to fetch terms for a specific Term Set directly from SharePoint.
   * In SharePoint Online REST v2.1, /terms returns root terms or flat terms with parent relationships.
   */
  private static async _fetchTermsForSet(
    targetUrl: string,
    setId: string,
    setName: string,
    groupName: string,
    headers: Record<string, string>
  ): Promise<ITermStoreTag[]> {
    try {
      // Query terms with relations expanded so parent/child links are explicitly available
      const termsEndpoint = `${targetUrl}/_api/v2.1/termStore/sets/${setId}/terms?$expand=relations,children`;
      const resp = await fetch(termsEndpoint, { headers, credentials: 'include' });
      if (resp.ok) {
        const data = await resp.json();
        if (data && Array.isArray(data.value)) {
          return TaxonomyService._parseTermsArray(data.value, setId, setName, groupName);
        }
      }
    } catch (err) {
      console.warn(`[TaxonomyService] Failed to query terms for set ${setName}:`, err);
    }
    return [];
  }

  /**
   * Parses SharePoint v2.1 terms payload into strongly typed ITermStoreTag objects with synonyms.
   * Preserves parentId from parent property, relations, or parentTermId.
   */
  private static _parseTermsArray(
    rawTerms: Array<{
      id: string;
      labels?: Array<{ name: string; isDefault?: boolean; languageTag?: string }>;
      children?: Array<any>;
      parent?: { id: string };
      relations?: Array<{ relationType?: string; type?: string; targetId?: string; fromTerm?: { id: string }; toTerm?: { id: string } }>;
    }>,
    setId: string,
    setName: string,
    groupName: string,
    parentPath?: string,
    parentTermId?: string
  ): ITermStoreTag[] {
    const rawList: ITermStoreTag[] = [];

    rawTerms.forEach((t) => {
      let primaryLabel = 'Term';
      const synonyms: string[] = [];

      if (Array.isArray(t.labels) && t.labels.length > 0) {
        const defaultLabelObj = t.labels.find((l) => l.isDefault === true);
        primaryLabel = defaultLabelObj?.name || t.labels[0].name;

        t.labels.forEach((l) => {
          if (l.name && l.name !== primaryLabel && synonyms.indexOf(l.name) === -1) {
            synonyms.push(l.name);
          }
        });
      }

      const currentPath = parentPath
        ? `${parentPath} > ${primaryLabel}`
        : `${groupName} > ${setName} > ${primaryLabel}`;

      // Extract parent ID from t.parent, relations array, or passed parentTermId
      let directParentId = t.parent?.id || parentTermId;
      if (!directParentId && Array.isArray(t.relations)) {
        for (const rel of t.relations) {
          // In v2.1 relations: relationType 'parent' or targetId or fromTerm
          if (rel.relationType === 'parent' || rel.type === 'parent') {
            directParentId = rel.targetId || rel.fromTerm?.id || rel.toTerm?.id;
            break;
          }
        }
      }

      const parsedTerm: ITermStoreTag = {
        id: t.id,
        label: primaryLabel,
        termSetId: setId,
        termSetName: setName,
        path: currentPath,
        synonyms: synonyms.length > 0 ? synonyms : undefined,
        parentId: directParentId,
        children: []
      };

      // Parse nested child terms if returned inline
      if (Array.isArray(t.children) && t.children.length > 0) {
        parsedTerm.children = TaxonomyService._parseTermsArray(t.children, setId, setName, groupName, currentPath, t.id);
      }

      rawList.push(parsedTerm);
    });

    // Assemble flat list into a hierarchical tree based on parentId and relations
    return TaxonomyService.organizeTermsIntoTree(rawList);
  }

  /**
   * Organises a list of terms into a hierarchical tree based on parentId or path structure.
   */
  public static organizeTermsIntoTree(terms: ITermStoreTag[]): ITermStoreTag[] {
    if (!terms || terms.length === 0) return [];

    // Check if terms are already hierarchical (i.e. roots with non-empty children)
    const hasInlineChildren = terms.some((t) => Array.isArray(t.children) && t.children.length > 0);
    const hasParentIds = terms.some((t) => !!t.parentId);

    if (hasInlineChildren && !hasParentIds) {
      return terms;
    }

    const termMap = new Map<string, ITermStoreTag>();
    terms.forEach((t) => {
      // Clone so we don't mutate references unexpectedly
      termMap.set(t.id, {
        ...t,
        children: Array.isArray(t.children) ? [...t.children] : []
      });
    });

    const rootTerms: ITermStoreTag[] = [];

    // First pass: Link child to parent by parentId
    if (hasParentIds) {
      termMap.forEach((term) => {
        if (term.parentId && termMap.has(term.parentId)) {
          const parent = termMap.get(term.parentId)!;
          if (!parent.children) {
            parent.children = [];
          }
          // Avoid duplicate insertions
          if (!parent.children.some((c) => c.id === term.id)) {
            parent.children.push(term);
          }
        } else {
          rootTerms.push(term);
        }
      });
      return rootTerms;
    }

    // Second pass: If parentId is not present, check path depth (e.g. "Group > Set > Parent > Child")
    const termsByDepth = Array.from(termMap.values()).sort((a, b) => {
      const depthA = (a.path || '').split(' > ').length;
      const depthB = (b.path || '').split(' > ').length;
      return depthA - depthB;
    });

    // If all terms have the same depth (e.g. depth 3), return as-is
    const minDepth = termsByDepth.length > 0 ? (termsByDepth[0].path || '').split(' > ').length : 0;
    const maxDepth = termsByDepth.length > 0 ? (termsByDepth[termsByDepth.length - 1].path || '').split(' > ').length : 0;

    if (minDepth === maxDepth) {
      return termsByDepth;
    }

    // Link by matching path prefixes
    const assembledRoots: ITermStoreTag[] = [];
    termsByDepth.forEach((term) => {
      const segments = (term.path || '').split(' > ');
      if (segments.length <= minDepth) {
        assembledRoots.push(term);
      } else {
        // Find immediate parent by path (all segments except last)
        const parentPath = segments.slice(0, -1).join(' > ');
        const parentTerm = Array.from(termMap.values()).find((p) => p.path === parentPath);
        if (parentTerm) {
          if (!parentTerm.children) {
            parentTerm.children = [];
          }
          if (!parentTerm.children.some((c) => c.id === term.id)) {
            parentTerm.children.push(term);
          }
        } else {
          assembledRoots.push(term);
        }
      }
    });

    return assembledRoots;
  }

  /**
   * Returns all hierarchical term groups and their term sets directly from live SharePoint.
   */
  public static async getTermGroups(siteUrl?: string): Promise<ITermGroup[]> {
    await this.initializeFromSharePoint(siteUrl);
    return this._cachedGroups;
  }

  /**
   * Helper to recursively collect all terms and their descendants into a flat list.
   */
  public static flattenTerms(terms: ITermStoreTag[]): ITermStoreTag[] {
    const list: ITermStoreTag[] = [];
    const recurse = (arr: ITermStoreTag[]): void => {
      arr.forEach((t) => {
        list.push(t);
        if (Array.isArray(t.children) && t.children.length > 0) {
          recurse(t.children);
        }
      });
    };
    recurse(terms);
    return list;
  }

  /**
   * Fetches flattened list of terms matching optional search query and term set filter.
   */
  public static async getTerms(filterQuery?: string, termSetFilter?: string, siteUrl?: string): Promise<ITermStoreTag[]> {
    await this.initializeFromSharePoint(siteUrl);

    const allTerms: ITermStoreTag[] = [];
    const seenIds = new Set<string>();

    this._cachedGroups.forEach((g) => {
      g.termSets.forEach((s) => {
        const flatSetTerms = TaxonomyService.flattenTerms(s.terms);
        flatSetTerms.forEach((t) => {
          if (!seenIds.has(t.id)) {
            seenIds.add(t.id);
            allTerms.push(t);
          }
        });
      });
    });

    let results = allTerms;

    if (termSetFilter && termSetFilter !== 'all') {
      const targetFilter = termSetFilter.toLowerCase().trim();
      results = results.filter(
        (t) => t.termSetName?.toLowerCase() === targetFilter || (t.path && t.path.toLowerCase().indexOf(targetFilter) !== -1)
      );
    }

    if (!filterQuery || !filterQuery.trim()) {
      return results;
    }

    const q = filterQuery.toLowerCase().trim();
    return results.filter((t) => {
      const labelMatch = t.label.toLowerCase().indexOf(q) !== -1;
      const setMatch = t.termSetName ? t.termSetName.toLowerCase().indexOf(q) !== -1 : false;
      const pathMatch = t.path ? t.path.toLowerCase().indexOf(q) !== -1 : false;
      const synMatch = Array.isArray(t.synonyms) && t.synonyms.some((s) => s.toLowerCase().indexOf(q) !== -1);
      return labelMatch || setMatch || pathMatch || synMatch;
    });
  }

  /**
   * Returns list of unique term sets available across all live term groups.
   */
  public static getAvailableTermSets(): string[] {
    const sets = new Set<string>();
    this._cachedGroups.forEach((g) => {
      g.termSets.forEach((s) => {
        sets.add(s.name);
      });
    });
    return Array.from(sets);
  }

  /**
   * Fetches all terms in a given term set by name (case-insensitive) or set ID (including all descendants).
   */
  public static async getTermsByTermSet(termSetName: string, siteUrl?: string): Promise<ITermStoreTag[]> {
    await this.initializeFromSharePoint(siteUrl);
    const target = termSetName.toLowerCase().trim();
    const matches: ITermStoreTag[] = [];

    this._cachedGroups.forEach((g) => {
      g.termSets.forEach((s) => {
        if (s.name.toLowerCase() === target || s.id.toLowerCase() === target) {
          matches.push(...TaxonomyService.flattenTerms(s.terms));
        }
      });
    });

    return matches;
  }

  /**
   * Adds a new custom term to the active term group/set in-memory.
   */
  public static addTerm(label: string, termSetName: string = 'Our Sectors', groupName: string = 'Intranet'): ITermStoreTag {
    const trimmed = label.trim();
    const newTerm: ITermStoreTag = {
      id: `t-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      label: trimmed,
      termSetName: termSetName,
      path: `${groupName} > ${termSetName} > ${trimmed}`
    };

    let group = this._cachedGroups.find((g) => g.name.toLowerCase() === groupName.toLowerCase());
    if (!group) {
      group = { id: `grp-${Date.now()}`, name: groupName, termSets: [] };
      this._cachedGroups.push(group);
    }

    let set = group.termSets.find((s) => s.name.toLowerCase() === termSetName.toLowerCase());
    if (!set) {
      set = { id: `set-${Date.now()}`, name: termSetName, groupId: group.id, groupName: group.name, terms: [] };
      group.termSets.push(set);
    }

    set.terms.push(newTerm);
    return newTerm;
  }

  /**
   * Resolves a user profile raw value against either:
   * 1. A Term Set's terms & synonyms
   * 2. A specific Parent Term (and all its children/descendants) & synonyms
   * If a match is found, returns the canonical term label. Otherwise returns raw value.
   */
  public static async resolveTermFromSynonym(rawValue: string, termSetNameOrTermId: string, siteUrl?: string): Promise<string> {
    if (!rawValue || !rawValue.trim() || !termSetNameOrTermId || !termSetNameOrTermId.trim()) {
      return rawValue;
    }

    const trimmedVal = rawValue.trim().toLowerCase();
    await this.initializeFromSharePoint(siteUrl);

    const targetKey = termSetNameOrTermId.toLowerCase().trim();

    // Iterate through all cached groups and term sets
    for (const group of this._cachedGroups) {
      for (const set of group.termSets) {
        const allFlatTerms = TaxonomyService.flattenTerms(set.terms);

        // Case A: The linked entity is the entire Term Set
        if (set.name.toLowerCase() === targetKey || set.id.toLowerCase() === targetKey) {
          for (const term of allFlatTerms) {
            if (term.label && term.label.toLowerCase() === trimmedVal) {
              return term.label;
            }
            if (Array.isArray(term.synonyms)) {
              for (const synonym of term.synonyms) {
                if (synonym && synonym.trim().toLowerCase() === trimmedVal) {
                  return term.label;
                }
              }
            }
          }
        }

        // Case B: The linked entity is a specific Term within the set (e.g. "Our teams")
        const matchingParent = allFlatTerms.find(
          (t) => t.id.toLowerCase() === targetKey || t.label.toLowerCase() === targetKey
        );

        if (matchingParent) {
          // Check the matched term itself
          if (matchingParent.label && matchingParent.label.toLowerCase() === trimmedVal) {
            return matchingParent.label;
          }
          if (Array.isArray(matchingParent.synonyms)) {
            for (const synonym of matchingParent.synonyms) {
              if (synonym && synonym.trim().toLowerCase() === trimmedVal) {
                return matchingParent.label;
              }
            }
          }

          // Check all descendant terms under this parent
          const descendants = TaxonomyService.flattenTerms(matchingParent.children || []);
          for (const desc of descendants) {
            if (desc.label && desc.label.toLowerCase() === trimmedVal) {
              return desc.label;
            }
            if (Array.isArray(desc.synonyms)) {
              for (const synonym of desc.synonyms) {
                if (synonym && synonym.trim().toLowerCase() === trimmedVal) {
                  return desc.label;
                }
              }
            }
          }

          // Fallback: check all flat terms in set whose path starts with this term
          const parentPrefix = `${matchingParent.path} >`.toLowerCase();
          const pathChildren = allFlatTerms.filter((t) => t.path && t.path.toLowerCase().startsWith(parentPrefix));
          for (const pc of pathChildren) {
            if (pc.label && pc.label.toLowerCase() === trimmedVal) {
              return pc.label;
            }
            if (Array.isArray(pc.synonyms)) {
              for (const synonym of pc.synonyms) {
                if (synonym && synonym.trim().toLowerCase() === trimmedVal) {
                  return pc.label;
                }
              }
            }
          }
        }
      }
    }

    return rawValue;
  }
}
