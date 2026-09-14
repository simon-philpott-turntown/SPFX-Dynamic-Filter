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
        const groupsEndpoint = `${targetUrl}/_api/v2.1/termStore/groups?$expand=sets($expand=terms)`;
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
            terms?: Array<{
              id: string;
              labels?: Array<{ name: string; isDefault?: boolean; languageTag?: string }>;
            }>;
          }>;
        }>) {
          const groupName = g.displayName || g.name || 'Group';
          let termSets: ITermSet[] = [];

          // If sets came pre-expanded
          if (Array.isArray(g.sets) && g.sets.length > 0) {
            for (const s of g.sets) {
              const setName = s.localizedNames && s.localizedNames[0] ? s.localizedNames[0].name : 'Term Set';
              let terms: ITermStoreTag[] = [];

              if (Array.isArray(s.terms) && s.terms.length > 0) {
                terms = TaxonomyService._parseTermsArray(s.terms, s.id, setName, groupName);
              } else {
                // Query terms for this set individually
                terms = await TaxonomyService._fetchTermsForSet(targetUrl, s.id, setName, groupName, headers);
              }

              termSets.push({
                id: s.id,
                name: setName,
                groupId: g.id,
                groupName: groupName,
                terms: terms
              });
            }
          } else {
            // Query sets for this group directly if not expanded inline
            try {
              const setsEndpoint = `${targetUrl}/_api/v2.1/termStore/groups/${g.id}/sets`;
              const setsResp = await fetch(setsEndpoint, { headers, credentials: 'include' });
              if (setsResp.ok) {
                const setsData = await setsResp.json();
                if (setsData && Array.isArray(setsData.value)) {
                  for (const s of setsData.value as Array<{ id: string; localizedNames?: Array<{ name: string }> }>) {
                    const setName = s.localizedNames && s.localizedNames[0] ? s.localizedNames[0].name : 'Term Set';
                    const terms = await TaxonomyService._fetchTermsForSet(targetUrl, s.id, setName, groupName, headers);
                    termSets.push({
                      id: s.id,
                      name: setName,
                      groupId: g.id,
                      groupName: groupName,
                      terms: terms
                    });
                  }
                }
              }
            } catch (setErr) {
              console.warn(`[TaxonomyService] Error fetching sets for group ${groupName}:`, setErr);
            }
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
   */
  private static async _fetchTermsForSet(
    targetUrl: string,
    setId: string,
    setName: string,
    groupName: string,
    headers: Record<string, string>
  ): Promise<ITermStoreTag[]> {
    try {
      const termsEndpoint = `${targetUrl}/_api/v2.1/termStore/sets/${setId}/terms`;
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
   * Preserves parentId and nested children structure.
   */
  private static _parseTermsArray(
    rawTerms: Array<{
      id: string;
      labels?: Array<{ name: string; isDefault?: boolean; languageTag?: string }>;
      children?: Array<any>;
      parent?: { id: string };
      relations?: Array<{ type: string; id: string }>;
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

      const directParentId = t.parent?.id || parentTermId;

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

    // If terms came in as a flat list with parent references, assemble them into a nested tree
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

    // 1. Try finding terms by Term Set name or id
    const setTerms = await this.getTermsByTermSet(termSetNameOrTermId, siteUrl);
    if (setTerms.length > 0) {
      for (const term of setTerms) {
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

    // 2. Also search all terms across groups for parent term matches or child terms
    const targetKey = termSetNameOrTermId.toLowerCase().trim();
    for (const group of this._cachedGroups) {
      for (const set of group.termSets) {
        // If the selected entity matches the set name directly:
        if (set.name.toLowerCase() === targetKey || set.id.toLowerCase() === targetKey) {
          for (const term of set.terms) {
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

        // If the selected entity matches a parent term in the set:
        const matchingParentTerm = set.terms.find(
          (t) => t.id.toLowerCase() === targetKey || t.label.toLowerCase() === targetKey
        );

        if (matchingParentTerm) {
          // Check the parent term itself
          if (matchingParentTerm.label.toLowerCase() === trimmedVal) {
            return matchingParentTerm.label;
          }
          if (Array.isArray(matchingParentTerm.synonyms)) {
            for (const synonym of matchingParentTerm.synonyms) {
              if (synonym && synonym.trim().toLowerCase() === trimmedVal) {
                return matchingParentTerm.label;
              }
            }
          }

          // Check child terms that have this term's label or path in their path
          const parentPrefix = `${matchingParentTerm.path} >`.toLowerCase();
          const children = set.terms.filter(
            (t) => t.path && t.path.toLowerCase().startsWith(parentPrefix)
          );

          for (const child of children) {
            if (child.label.toLowerCase() === trimmedVal) {
              return child.label;
            }
            if (Array.isArray(child.synonyms)) {
              for (const synonym of child.synonyms) {
                if (synonym && synonym.trim().toLowerCase() === trimmedVal) {
                  return child.label;
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
