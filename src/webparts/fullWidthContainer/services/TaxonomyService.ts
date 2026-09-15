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
   * Primary target Term Set GUID for 'Our business' in Turner & Townsend Term Store taxonomy lookup.
   * Scopes queries directly to avoid querying unneeded groups and prevents 404 child term errors.
   */
  public static readonly OUR_BUSINESS_SET_ID: string = '930372f6-3b00-46f8-8e63-4b40edb8fd22';

  /**
   * Loaded live groups from SharePoint Online.
   * Initialised empty: strictly populated from live SharePoint Term Store.
   */
  private static _cachedGroups: ITermGroup[] = [];

  private static _isFetched: boolean = false;
  private static _fetchPromise: Promise<void> | null = null;

  /**
   * Queries the live SharePoint Online Term Store v2.1 REST API.
   * 1. Targets the designated 'Our business' Term Set:
   *    `/_api/v2.1/termStore/termSets/930372f6-3b00-46f8-8e63-4b40edb8fd22/getlegacychildren`
   * 2. For each root category (AI, Content categories, Our capabilities, Our regions, Our sectors, Our segments, Our teams),
   *    if childrenCount > 0, recursively calls:
   *    `/_api/v2.1/termStore/termSets/930372f6-3b00-46f8-8e63-4b40edb8fd22/terms/{termId}/getlegacychildren`
   * 3. Completely eliminates leaf-term 404 errors and preserves multi-tier hierarchy.
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

        const liveGroups: ITermGroup[] = [];

        // Fetch root terms of 'Our business' set using getlegacychildren
        const rootCategories = await TaxonomyService._fetchLegacyChildren(
          targetUrl,
          TaxonomyService.OUR_BUSINESS_SET_ID,
          'Our business',
          undefined,
          headers
        );

        if (rootCategories && rootCategories.length > 0) {
          // Represent each Tier 1 category (Our teams, Our sectors, Our capabilities, etc.)
          // as a distinct selectable Term Set within the 'Our business' group
          const termSets: ITermSet[] = rootCategories.map((cat) => ({
            id: cat.id,
            name: cat.label,
            groupId: TaxonomyService.OUR_BUSINESS_SET_ID,
            groupName: 'Our business',
            terms: cat.children && cat.children.length > 0 ? cat.children : [cat]
          }));

          // Also include the top-level 'Our business' set containing all root categories
          termSets.unshift({
            id: TaxonomyService.OUR_BUSINESS_SET_ID,
            name: 'Our business',
            groupId: TaxonomyService.OUR_BUSINESS_SET_ID,
            groupName: 'Our business',
            terms: rootCategories
          });

          liveGroups.push({
            id: TaxonomyService.OUR_BUSINESS_SET_ID,
            name: 'Our business',
            termSets: termSets
          });
        }

        // Fallback: If Our business set failed or returned no categories, fallback to generic /termSets query
        if (liveGroups.length === 0) {
          try {
            const fallbackSetsUrl = `${targetUrl}/_api/v2.1/termStore/termSets/${TaxonomyService.OUR_BUSINESS_SET_ID}`;
            const fbResp = await fetch(fallbackSetsUrl, { headers, credentials: 'include' });
            if (fbResp.ok) {
              const fbData = await fbResp.json();
              const setName = fbData.localizedNames && fbData.localizedNames[0] ? fbData.localizedNames[0].name : (fbData.name || 'Our business');
              liveGroups.push({
                id: TaxonomyService.OUR_BUSINESS_SET_ID,
                name: 'Our business',
                termSets: [{
                  id: TaxonomyService.OUR_BUSINESS_SET_ID,
                  name: setName,
                  groupId: TaxonomyService.OUR_BUSINESS_SET_ID,
                  groupName: 'Our business',
                  terms: rootCategories
                }]
              });
            }
          } catch (fbErr) {
            console.warn('[TaxonomyService] Fallback set fetch error:', fbErr);
          }
        }

        // Assign live groups returned from the tenant Term Store
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
   * Recursively fetches child terms for a set or term using SharePoint v2.1 getlegacychildren endpoint.
   * Only calls child endpoints if childrenCount > 0, completely avoiding 404 Not Found errors on leaf nodes.
   */
  private static async _fetchLegacyChildren(
    targetUrl: string,
    setId: string,
    setName: string,
    parentTermId: string | undefined,
    headers: Record<string, string>,
    parentPath?: string
  ): Promise<ITermStoreTag[]> {
    try {
      const endpoint = parentTermId
        ? `${targetUrl}/_api/v2.1/termStore/termSets/${setId}/terms/${parentTermId}/getlegacychildren`
        : `${targetUrl}/_api/v2.1/termStore/termSets/${setId}/getlegacychildren`;

      const resp = await fetch(endpoint, { headers, credentials: 'include' });
      if (!resp.ok) {
        return [];
      }

      const data = await resp.json();
      if (!data || !Array.isArray(data.value) || data.value.length === 0) {
        return [];
      }

      const tags: ITermStoreTag[] = [];

      for (const t of data.value) {
        const normId = TaxonomyService._normId(t.id);
        let primaryLabel = 'Term';
        const synonyms: string[] = [];

        if (Array.isArray(t.labels) && t.labels.length > 0) {
          const defaultLabelObj = t.labels.find((l: any) => l.isDefault === true);
          primaryLabel = defaultLabelObj?.name || t.labels[0].name;

          t.labels.forEach((l: any) => {
            if (l.name && l.name !== primaryLabel && synonyms.indexOf(l.name) === -1) {
              synonyms.push(l.name);
            }
          });
        }

        const currentPath = parentPath
          ? `${parentPath} > ${primaryLabel}`
          : `${setName} > ${primaryLabel}`;

        const termTag: ITermStoreTag = {
          id: normId,
          label: primaryLabel,
          termSetId: setId,
          termSetName: setName,
          path: currentPath,
          synonyms: synonyms.length > 0 ? synonyms : undefined,
          parentId: parentTermId ? TaxonomyService._normId(parentTermId) : undefined,
          children: []
        };

        // Recursively fetch subterms ONLY if childrenCount > 0
        const childrenCount = typeof t.childrenCount === 'number' ? t.childrenCount : 0;
        if (childrenCount > 0) {
          try {
            const nestedChildren = await TaxonomyService._fetchLegacyChildren(
              targetUrl,
              setId,
              setName,
              t.id,
              headers,
              currentPath
            );
            termTag.children = nestedChildren;
          } catch (childErr) {
            console.warn(`[TaxonomyService] Error querying children for ${primaryLabel}:`, childErr);
          }
        }

        tags.push(termTag);
      }

      return tags;
    } catch (err) {
      console.warn(`[TaxonomyService] Error in _fetchLegacyChildren for parent ${parentTermId || 'root'}:`, err);
      return [];
    }
  }

  /**
   * Parses SharePoint v2.1 terms payload into strongly typed ITermStoreTag objects with synonyms.
   * Preserves parentId from parent property, relations map, or parentTermId.
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
    parentTermId?: string,
    parentIdMap?: Map<string, string>,
    rootTermIds?: Set<string>
  ): ITermStoreTag[] {
    const rawList: ITermStoreTag[] = [];

    rawTerms.forEach((t) => {
      const normId = TaxonomyService._normId(t.id);
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

      // Extract parent ID from:
      // 1. parentIdMap derived from /relations
      // 2. t.parent.id
      // 3. inline t.relations
      // 4. parentTermId passed down
      let directParentId = (parentIdMap && normId ? parentIdMap.get(normId) : undefined) || t.parent?.id || parentTermId;
      if (!directParentId && Array.isArray(t.relations)) {
        for (const rel of t.relations) {
          if (rel.relationType === 'parent' || rel.type === 'parent') {
            directParentId = rel.targetId || rel.fromTerm?.id || rel.toTerm?.id;
            break;
          }
        }
      }

      // If this term is confirmed to be a root term via /children, it has NO parent
      if (rootTermIds && rootTermIds.has(normId)) {
        directParentId = undefined;
      }

      const parsedTerm: ITermStoreTag = {
        id: normId,
        label: primaryLabel,
        termSetId: setId,
        termSetName: setName,
        path: currentPath,
        synonyms: synonyms.length > 0 ? synonyms : undefined,
        parentId: directParentId ? TaxonomyService._normId(directParentId) : undefined,
        children: []
      };

      // Parse nested child terms if returned inline
      if (Array.isArray(t.children) && t.children.length > 0) {
        parsedTerm.children = TaxonomyService._parseTermsArray(
          t.children,
          setId,
          setName,
          groupName,
          currentPath,
          t.id,
          parentIdMap,
          rootTermIds
        );
      }

      rawList.push(parsedTerm);
    });

    // Assemble flat list into a hierarchical tree based on parentId and relations
    return TaxonomyService.organizeTermsIntoTree(rawList, rootTermIds);
  }

  /**
   * Helper to normalize GUID strings (strip brackets and lowercase)
   */
  private static _normId(id?: string): string {
    return (id || '').toLowerCase().replace(/[{}]/g, '').trim();
  }

  /**
   * Organises a list of terms into a hierarchical tree based on parentId or path structure.
   */
  public static organizeTermsIntoTree(terms: ITermStoreTag[], rootTermIds?: Set<string>): ITermStoreTag[] {
    if (!terms || terms.length === 0) return [];

    // Map by normalized ID
    const termMap = new Map<string, ITermStoreTag>();
    terms.forEach((t) => {
      const normId = TaxonomyService._normId(t.id);
      termMap.set(normId, {
        ...t,
        id: normId,
        parentId: t.parentId ? TaxonomyService._normId(t.parentId) : undefined,
        children: Array.isArray(t.children) ? [...t.children] : []
      });
    });

    const hasParentIds = Array.from(termMap.values()).some((t) => !!t.parentId);

    if (hasParentIds) {
      const rootTerms: ITermStoreTag[] = [];
      termMap.forEach((term) => {
        // If explicitly in rootTermIds, it is a root
        const isExplicitRoot = rootTermIds && rootTermIds.has(term.id);

        if (!isExplicitRoot && term.parentId && termMap.has(term.parentId) && term.parentId !== term.id) {
          const parent = termMap.get(term.parentId)!;
          if (!parent.children) {
            parent.children = [];
          }
          if (!parent.children.some((c) => TaxonomyService._normId(c.id) === term.id)) {
            parent.children.push(term);
          }
        } else {
          rootTerms.push(term);
        }
      });
      return rootTerms;
    }

    // Second pass: If parentId is not populated by API, assemble by path segments
    const termsByDepth = Array.from(termMap.values()).sort((a, b) => {
      const depthA = (a.path || '').split(' > ').length;
      const depthB = (b.path || '').split(' > ').length;
      return depthA - depthB;
    });

    const minDepth = termsByDepth.length > 0 ? (termsByDepth[0].path || '').split(' > ').length : 0;
    const maxDepth = termsByDepth.length > 0 ? (termsByDepth[termsByDepth.length - 1].path || '').split(' > ').length : 0;

    if (minDepth === maxDepth) {
      return termsByDepth;
    }

    const assembledRoots: ITermStoreTag[] = [];
    termsByDepth.forEach((term) => {
      const segments = (term.path || '').split(' > ');
      if (segments.length <= minDepth) {
        assembledRoots.push(term);
      } else {
        const parentPath = segments.slice(0, -1).join(' > ');
        const parentTerm = Array.from(termMap.values()).find((p) => p.path === parentPath);
        if (parentTerm) {
          if (!parentTerm.children) {
            parentTerm.children = [];
          }
          if (!parentTerm.children.some((c) => TaxonomyService._normId(c.id) === term.id)) {
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
  public static async getTermGroups(siteUrl?: string, forceRefresh: boolean = false): Promise<ITermGroup[]> {
    await this.initializeFromSharePoint(siteUrl, forceRefresh);
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
  public static async getTerms(filterQuery?: string, termSetFilter?: string, siteUrl?: string, forceRefresh: boolean = false): Promise<ITermStoreTag[]> {
    await this.initializeFromSharePoint(siteUrl, forceRefresh);

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
    // Also prepare normalized slash version (e.g. forward slash vs backslash)
    const normSlashVal = trimmedVal.replace(/\\/g, '/');

    await this.initializeFromSharePoint(siteUrl);

    const targetKey = termSetNameOrTermId.toLowerCase().trim();
    const targetNormId = TaxonomyService._normId(termSetNameOrTermId);

    // Helper to check if a term matches the value either by its canonical label or any of its synonyms
    const checkTermMatch = (term: ITermStoreTag): boolean => {
      if (term.label) {
        const lblLower = term.label.toLowerCase().trim();
        if (lblLower === trimmedVal || lblLower.replace(/\\/g, '/') === normSlashVal) {
          return true;
        }
      }
      if (Array.isArray(term.synonyms)) {
        for (const synonym of term.synonyms) {
          if (synonym) {
            const synLower = synonym.trim().toLowerCase();
            if (synLower === trimmedVal || synLower.replace(/\\/g, '/') === normSlashVal) {
              return true;
            }
          }
        }
      }
      return false;
    };

    // Iterate through all cached groups and term sets
    for (const group of this._cachedGroups) {
      for (const set of group.termSets) {
        const allFlatTerms = TaxonomyService.flattenTerms(set.terms);

        // Case A: The linked entity is the entire Term Set (by name or ID)
        if (
          set.name.toLowerCase() === targetKey ||
          set.id.toLowerCase() === targetKey ||
          TaxonomyService._normId(set.id) === targetNormId
        ) {
          for (const term of allFlatTerms) {
            if (checkTermMatch(term)) {
              return term.label;
            }
          }
        }

        // Case B: The linked entity is a specific Term (by name or ID, e.g. "Our teams")
        const matchingParent = allFlatTerms.find(
          (t) =>
            t.label.toLowerCase().trim() === targetKey ||
            t.id.toLowerCase() === targetKey ||
            TaxonomyService._normId(t.id) === targetNormId
        );

        if (matchingParent) {
          // Check matching parent itself
          if (checkTermMatch(matchingParent)) {
            return matchingParent.label;
          }

          // Check all nested descendants under this parent term
          const descendants = TaxonomyService.flattenTerms(matchingParent.children || []);
          for (const desc of descendants) {
            if (checkTermMatch(desc)) {
              return desc.label;
            }
          }

          // Fallback: check all flat terms in set whose path contains or starts with this parent
          const parentPrefix = `${matchingParent.path || matchingParent.label} >`.toLowerCase();
          const pathChildren = allFlatTerms.filter(
            (t) => t.path && t.path.toLowerCase().indexOf(parentPrefix) !== -1
          );
          for (const pc of pathChildren) {
            if (checkTermMatch(pc)) {
              return pc.label;
            }
          }
        }
      }
    }

    // Case C: Global fallback across all terms in all sets if termSetNameOrTermId matches any term or set anywhere
    for (const group of this._cachedGroups) {
      for (const set of group.termSets) {
        const allFlatTerms = TaxonomyService.flattenTerms(set.terms);
        for (const term of allFlatTerms) {
          if (checkTermMatch(term)) {
            // If the term's path or termSetName contains the targetKey, prioritize it
            if (
              (term.path && term.path.toLowerCase().indexOf(targetKey) !== -1) ||
              (term.termSetName && term.termSetName.toLowerCase() === targetKey)
            ) {
              return term.label;
            }
          }
        }
      }
    }

    return rawValue;
  }
}
