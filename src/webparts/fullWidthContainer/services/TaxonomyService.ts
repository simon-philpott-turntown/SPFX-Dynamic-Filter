/**
 * @file TaxonomyService.ts
 * @description Service for querying SharePoint Global Term Store taxonomy directly in SharePoint Online (SPFx)
 * with hierarchical term groups, term sets, live REST/v2.1 endpoints, and comprehensive tenant fallback cache.
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
   * Comprehensive default tenant taxonomy structure mirroring SharePoint Online Term Store.
   * Note: In tenant Term Store, "Intranet" and "AMCL Terms" are top-level Global Term Groups.
   * "Our business" is the primary Term Set under the "Intranet" Group (and also aliases child sets).
   */
  private static _cachedGroups: ITermGroup[] = [
    {
      id: 'grp-amcl',
      name: 'AMCL Terms',
      termSets: [
        {
          id: 'set-amcl-core',
          name: 'Our teams',
          groupId: 'grp-amcl',
          groupName: 'AMCL Terms',
          terms: [
            {
              id: 'team-km-amcl',
              label: 'Knowledge Management',
              termSetName: 'Our teams',
              path: 'AMCL Terms > Our teams > Knowledge Management',
              synonyms: [
                'TT Company\\Support Services\\KM',
                'KM',
                'Knowledge Services',
                'Support Services\\KM'
              ]
            },
            {
              id: 'team-ai-amcl',
              label: 'AI & Data Intelligence',
              termSetName: 'Our teams',
              path: 'AMCL Terms > Our teams > AI & Data Intelligence',
              synonyms: [
                'AI',
                'Artificial Intelligence',
                'Data Science'
              ]
            },
            {
              id: 'team-advisory-amcl',
              label: 'Programme Advisory',
              termSetName: 'Our teams',
              path: 'AMCL Terms > Our teams > Programme Advisory',
              synonyms: [
                'Advisory',
                'Consulting',
                'Strategic Advisory'
              ]
            }
          ]
        },
        {
          id: 'set-amcl-capabilities',
          name: 'Core Capabilities',
          groupId: 'grp-amcl',
          groupName: 'AMCL Terms',
          terms: [
            { id: 'amcl-frameworks', label: 'Asset management frameworks', termSetName: 'Core Capabilities', path: 'AMCL Terms > Core Capabilities > Asset management frameworks' },
            { id: 'amcl-iso55000', label: 'ISO 55000 compliance', termSetName: 'Core Capabilities', path: 'AMCL Terms > Core Capabilities > ISO 55000 compliance' },
            { id: 'amcl-readiness', label: 'Operational readiness', termSetName: 'Core Capabilities', path: 'AMCL Terms > Core Capabilities > Operational readiness' }
          ]
        }
      ]
    },
    {
      id: 'grp-intranet',
      name: 'Intranet',
      termSets: [
        {
          id: 'set-our-business',
          name: 'Our business',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            // Level 1 children of Our business as terms
            { id: 'ob-our-teams', label: 'Our teams', termSetName: 'Our business', path: 'Intranet > Our business > Our teams' },
            {
              id: 'team-km',
              label: 'Knowledge Management',
              termSetName: 'Our business',
              path: 'Intranet > Our business > Our teams > Knowledge Management',
              synonyms: [
                'TT Company\\Support Services\\KM',
                'KM',
                'Knowledge Services',
                'Support Services\\KM'
              ]
            },
            {
              id: 'team-ai',
              label: 'AI & Data Intelligence',
              termSetName: 'Our business',
              path: 'Intranet > Our business > Our teams > AI & Data Intelligence',
              synonyms: [
                'AI',
                'Artificial Intelligence',
                'Data Science'
              ]
            },
            {
              id: 'team-advisory',
              label: 'Programme Advisory',
              termSetName: 'Our business',
              path: 'Intranet > Our business > Our teams > Programme Advisory',
              synonyms: [
                'Advisory',
                'Consulting',
                'Strategic Advisory'
              ]
            },
            { id: 'ob-ai', label: 'AI', termSetName: 'Our business', path: 'Intranet > Our business > AI' },
            { id: 'ob-content-categories', label: 'Content categories', termSetName: 'Our business', path: 'Intranet > Our business > Content categories' },
            { id: 'ob-our-capabilities', label: 'Our capabilities', termSetName: 'Our business', path: 'Intranet > Our business > Our capabilities' },
            { id: 'ob-our-regions', label: 'Our regions', termSetName: 'Our business', path: 'Intranet > Our business > Our regions' },
            { id: 'reg-americas', label: 'Americas', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Americas' },
            { id: 'reg-latam', label: 'Latin America', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Americas > Latin America' },
            { id: 'reg-argentina', label: 'Argentina', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Americas > Latin America > Argentina' },
            { id: 'reg-brazil', label: 'Brazil', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Americas > Latin America > Brazil' },
            { id: 'reg-chile', label: 'Chile', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Americas > Latin America > Chile' },
            { id: 'reg-uk', label: 'UK', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > UK' },
            { id: 'reg-europe', label: 'Europe', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Europe' },
            { id: 'reg-middle-east', label: 'Middle East', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Middle East' },
            { id: 'reg-asia-pacific', label: 'Asia Pacific', termSetName: 'Our business', path: 'Intranet > Our business > Our regions > Asia Pacific' }
          ]
        },
        {
          id: 'set-our-teams-direct',
          name: 'Our teams',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            {
              id: 'team-km-direct',
              label: 'Knowledge Management',
              termSetName: 'Our teams',
              path: 'Intranet > Our business > Our teams > Knowledge Management',
              synonyms: [
                'TT Company\\Support Services\\KM',
                'KM',
                'Knowledge Services',
                'Support Services\\KM'
              ]
            },
            {
              id: 'team-ai-direct',
              label: 'AI & Data Intelligence',
              termSetName: 'Our teams',
              path: 'Intranet > Our business > Our teams > AI & Data Intelligence',
              synonyms: [
                'AI',
                'Artificial Intelligence',
                'Data Science'
              ]
            },
            {
              id: 'team-advisory-direct',
              label: 'Programme Advisory',
              termSetName: 'Our teams',
              path: 'Intranet > Our business > Our teams > Programme Advisory',
              synonyms: [
                'Advisory',
                'Consulting',
                'Strategic Advisory'
              ]
            }
          ]
        },
        {
          id: 'set-our-capabilities',
          name: 'Our capabilities',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'cap-prog-adv', label: 'Programme advisory', termSetName: 'Our capabilities', path: 'Intranet > Our business > Our capabilities > Programme advisory' },
            { id: 'cap-cost-comm', label: 'Cost and commercial management', termSetName: 'Our capabilities', path: 'Intranet > Our business > Our capabilities > Cost and commercial management' },
            { id: 'cap-controls', label: 'Controls and performance', termSetName: 'Our capabilities', path: 'Intranet > Our business > Our capabilities > Controls and performance' },
            { id: 'cap-proj-mgmt', label: 'Project management', termSetName: 'Our capabilities', path: 'Intranet > Our business > Our capabilities > Project management' },
            { id: 'cap-proc-supply', label: 'Procurement and supply chain', termSetName: 'Our capabilities', path: 'Intranet > Our business > Our capabilities > Procurement and supply chain' },
            { id: 'cap-construction', label: 'Construction management', termSetName: 'Our capabilities', path: 'Intranet > Our business > Our capabilities > Construction management' },
            { id: 'cap-sustainability', label: 'Sustainability', termSetName: 'Our capabilities', path: 'Intranet > Our capabilities > Sustainability' },
            { id: 'cap-digital', label: 'Digital', termSetName: 'Our capabilities', path: 'Intranet > Our capabilities > Digital' },
            { id: 'cap-asset-consulting', label: 'Asset and building consultancy', termSetName: 'Our capabilities', path: 'Intranet > Our capabilities > Asset and building consultancy' }
          ]
        },
        {
          id: 'set-our-regions',
          name: 'Our regions',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'reg-uk-intranet', label: 'UK', termSetName: 'Our regions', path: 'Intranet > Our regions > UK' },
            { id: 'reg-europe-intranet', label: 'Europe', termSetName: 'Our regions', path: 'Intranet > Our regions > Europe' },
            { id: 'reg-north-america-intranet', label: 'North America', termSetName: 'Our regions', path: 'Intranet > Our regions > North America' },
            { id: 'reg-middle-east-intranet', label: 'Middle East', termSetName: 'Our regions', path: 'Intranet > Our regions > Middle East' },
            { id: 'reg-asia-pacific-intranet', label: 'Asia Pacific', termSetName: 'Our regions', path: 'Intranet > Our regions > Asia Pacific' },
            { id: 'reg-latin-america-intranet', label: 'Latin America', termSetName: 'Our regions', path: 'Intranet > Our regions > Latin America' },
            { id: 'reg-argentina-intranet', label: 'Argentina', termSetName: 'Our regions', path: 'Intranet > Our regions > Latin America > Argentina' },
            { id: 'reg-brazil-intranet', label: 'Brazil', termSetName: 'Our regions', path: 'Intranet > Our regions > Latin America > Brazil' },
            { id: 'reg-chile-intranet', label: 'Chile', termSetName: 'Our regions', path: 'Intranet > Our regions > Latin America > Chile' }
          ]
        },
        {
          id: 'set-content-categories',
          name: 'Content categories',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'cat-methodology', label: 'Service methodologies', termSetName: 'Content categories', path: 'Intranet > Our business > Content categories > Service methodologies' },
            { id: 'cat-frameworks', label: 'Frameworks and standards', termSetName: 'Content categories', path: 'Intranet > Our business > Content categories > Frameworks and standards' },
            { id: 'cat-case-studies', label: 'Case studies and insights', termSetName: 'Content categories', path: 'Intranet > Our business > Content categories > Case studies and insights' }
          ]
        },
        {
          id: 'set-ai',
          name: 'AI',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'ai-genai', label: 'Generative AI', termSetName: 'AI', path: 'Intranet > Our business > AI > Generative AI' },
            { id: 'ai-analytics', label: 'Predictive analytics', termSetName: 'AI', path: 'Intranet > Our business > AI > Predictive analytics' },
            { id: 'ai-automation', label: 'Intelligent automation', termSetName: 'AI', path: 'Intranet > Our business > AI > Intelligent automation' }
          ]
        },
        {
          id: 'set-our-sectors-legacy',
          name: 'Our Sectors',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'sec-clean-energy-leg', label: 'Clean energy', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Clean energy' },
            { id: 'sec-conv-power', label: 'Conventional and low carbon power', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Conventional and low carbon power' },
            { id: 'sec-data-centres-leg', label: 'Data centres', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Data centres' },
            { id: 'sec-defense', label: 'Defense', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Defense' },
            { id: 'sec-education-leg', label: 'Education', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Education' },
            { id: 'sec-elec-trans', label: 'Electrical transmission and distribution', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Electrical transmission and distribution' },
            { id: 'sec-energy-res-leg', label: 'Energy and resources', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Energy and resources' },
            { id: 'sec-environment', label: 'Environment', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Environment' },
            { id: 'sec-fin-prof', label: 'Finance and professional services', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Finance and professional services' },
            { id: 'sec-health-leg', label: 'Health', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Health' },
            { id: 'sec-infra-leg', label: 'Infrastructure', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Infrastructure' },
            { id: 'sec-prop', label: 'Real estate and property', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Real estate and property' },
            { id: 'sec-av', label: 'Aviation and transport', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Aviation and transport' },
            { id: 'sec-gov', label: 'Government and public sector', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Government and public sector' },
            { id: 'sec-tech', label: 'Technology and digital', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Technology and digital' },
            { id: 'sec-mining', label: 'Mining and metals', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Mining and metals' },
            { id: 'sec-life-sci', label: 'Life sciences and pharmaceuticals', termSetName: 'Our Sectors', path: 'Intranet > Our Sectors > Life sciences and pharmaceuticals' }
          ]
        },
        {
          id: 'set-our-segments-legacy',
          name: 'Our Segments',
          groupId: 'grp-intranet',
          groupName: 'Intranet',
          terms: [
            { id: 'seg-comm', label: 'Commercial advisory', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Commercial advisory' },
            { id: 'seg-prog', label: 'Programme management', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Programme management' },
            { id: 'seg-proj', label: 'Project controls', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Project controls' },
            { id: 'seg-cost', label: 'Cost and commercial management', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Cost and commercial management' },
            { id: 'seg-strat', label: 'Strategic advisory', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Strategic advisory' },
            { id: 'seg-esg', label: 'Sustainability, ESG and net zero', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Sustainability, ESG and net zero' },
            { id: 'seg-proc', label: 'Procurement and supply chain', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Procurement and supply chain' },
            { id: 'seg-disp', label: 'Dispute resolution', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Dispute resolution' },
            { id: 'seg-asset-ops', label: 'Asset management and operations', termSetName: 'Our Segments', path: 'Intranet > Our Segments > Asset management and operations' }
          ]
        }
      ]
    },
    {
      id: 'grp-our-business',
      name: 'Our business',
      termSets: [
        {
          id: 'set-our-teams-ob',
          name: 'Our teams',
          groupId: 'grp-our-business',
          groupName: 'Our business',
          terms: [
            {
              id: 'team-km-ob',
              label: 'Knowledge Management',
              termSetName: 'Our teams',
              path: 'Our business > Our teams > Knowledge Management',
              synonyms: [
                'TT Company\\Support Services\\KM',
                'KM',
                'Knowledge Services',
                'Support Services\\KM'
              ]
            },
            {
              id: 'team-ai-ob',
              label: 'AI & Data Intelligence',
              termSetName: 'Our teams',
              path: 'Our business > Our teams > AI & Data Intelligence',
              synonyms: [
                'AI',
                'Artificial Intelligence',
                'Data Science'
              ]
            },
            {
              id: 'team-advisory-ob',
              label: 'Programme Advisory',
              termSetName: 'Our teams',
              path: 'Our business > Our teams > Programme Advisory',
              synonyms: [
                'Advisory',
                'Consulting',
                'Strategic Advisory'
              ]
            }
          ]
        },
        {
          id: 'set-our-regions-ob',
          name: 'Our regions',
          groupId: 'grp-our-business',
          groupName: 'Our business',
          terms: [
            { id: 'reg-uk-ob', label: 'UK', termSetName: 'Our regions', path: 'Our business > Our regions > UK' },
            { id: 'reg-europe-ob', label: 'Europe', termSetName: 'Our regions', path: 'Our business > Our regions > Europe' },
            { id: 'reg-north-america-ob', label: 'North America', termSetName: 'Our regions', path: 'Our business > Our regions > North America' },
            { id: 'reg-middle-east-ob', label: 'Middle East', termSetName: 'Our regions', path: 'Our business > Our regions > Middle East' },
            { id: 'reg-asia-pacific-ob', label: 'Asia Pacific', termSetName: 'Our regions', path: 'Our business > Our regions > Asia Pacific' },
            { id: 'reg-latin-america-ob', label: 'Latin America', termSetName: 'Our regions', path: 'Our business > Our regions > Latin America' },
            { id: 'reg-argentina-ob', label: 'Argentina', termSetName: 'Our business > Our regions', path: 'Our business > Our regions > Latin America > Argentina' },
            { id: 'reg-brazil-ob', label: 'Brazil', termSetName: 'Our business > Our regions', path: 'Our business > Our regions > Latin America > Brazil' },
            { id: 'reg-chile-ob', label: 'Chile', termSetName: 'Our business > Our regions', path: 'Our business > Our regions > Latin America > Chile' }
          ]
        },
        {
          id: 'set-our-capabilities-ob',
          name: 'Our capabilities',
          groupId: 'grp-our-business',
          groupName: 'Our business',
          terms: [
            { id: 'cap-prog-adv-ob', label: 'Programme advisory', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Programme advisory' },
            { id: 'cap-cost-comm-ob', label: 'Cost and commercial management', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Cost and commercial management' },
            { id: 'cap-controls-ob', label: 'Controls and performance', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Controls and performance' },
            { id: 'cap-proj-mgmt-ob', label: 'Project management', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Project management' },
            { id: 'cap-proc-supply-ob', label: 'Procurement and supply chain', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Procurement and supply chain' },
            { id: 'cap-construction-ob', label: 'Construction management', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Construction management' },
            { id: 'cap-sustainability-ob', label: 'Sustainability', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Sustainability' },
            { id: 'cap-digital-ob', label: 'Digital', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Digital' },
            { id: 'cap-asset-consulting-ob', label: 'Asset and building consultancy', termSetName: 'Our capabilities', path: 'Our business > Our capabilities > Asset and building consultancy' }
          ]
        },
        {
          id: 'set-ai-ob',
          name: 'AI',
          groupId: 'grp-our-business',
          groupName: 'Our business',
          terms: [
            { id: 'ai-genai-ob', label: 'Generative AI', termSetName: 'AI', path: 'Our business > AI > Generative AI' },
            { id: 'ai-analytics-ob', label: 'Predictive analytics', termSetName: 'AI', path: 'Our business > AI > Predictive analytics' },
            { id: 'ai-automation-ob', label: 'Intelligent automation', termSetName: 'AI', path: 'Our business > AI > Intelligent automation' }
          ]
        },
        {
          id: 'set-content-categories-ob',
          name: 'Content categories',
          groupId: 'grp-our-business',
          groupName: 'Our business',
          terms: [
            { id: 'cat-methodology-ob', label: 'Service methodologies', termSetName: 'Content categories', path: 'Our business > Content categories > Service methodologies' },
            { id: 'cat-frameworks-ob', label: 'Frameworks and standards', termSetName: 'Content categories', path: 'Our business > Content categories > Frameworks and standards' },
            { id: 'cat-case-studies-ob', label: 'Case studies and insights', termSetName: 'Content categories', path: 'Our business > Content categories > Case studies and insights' }
          ]
        },
        {
          id: 'set-our-sectors-ob',
          name: 'Our sectors',
          groupId: 'grp-our-business',
          groupName: 'Our business',
          terms: [
            { id: 'sec-aviation', label: 'Aviation and transport', termSetName: 'Our sectors', path: 'Our business > Our sectors > Aviation and transport' },
            { id: 'sec-clean-energy', label: 'Clean energy', termSetName: 'Our sectors', path: 'Our business > Our sectors > Clean energy' },
            { id: 'sec-data-centres', label: 'Data centres', termSetName: 'Our sectors', path: 'Our business > Our sectors > Data centres' },
            { id: 'sec-health', label: 'Health', termSetName: 'Our sectors', path: 'Our business > Our sectors > Health' },
            { id: 'sec-education', label: 'Education', termSetName: 'Our sectors', path: 'Our business > Our sectors > Education' },
            { id: 'sec-utilities', label: 'Utilities and water', termSetName: 'Our sectors', path: 'Our business > Our sectors > Utilities and water' }
          ]
        },
        {
          id: 'set-our-segments-ob',
          name: 'Our segments',
          groupId: 'grp-our-business',
          groupName: 'Our business',
          terms: [
            { id: 'seg-energy-res', label: 'Energy & Natural Resources', termSetName: 'Our segments', path: 'Our business > Our segments > Energy & Natural Resources' },
            { id: 'seg-infra', label: 'Infrastructure', termSetName: 'Our segments', path: 'Our business > Our segments > Infrastructure' },
            { id: 'seg-real-estate', label: 'Real estate', termSetName: 'Our segments', path: 'Our business > Our segments > Real estate' },
            { id: 'seg-defence', label: 'Defense & security', termSetName: 'Our segments', path: 'Our business > Our segments > Defense & security' }
          ]
        }
      ]
    }
  ];

  private static _isFetched: boolean = false;

  /**
   * Queries the live SharePoint Online Term Store v2.1 REST API if available on the current site context.
   * Merges live data with fallback data so that groups/term sets with unexpanded REST sets or empty
   * responses are never rendered as empty.
   */
  public static async initializeFromSharePoint(siteUrl?: string): Promise<void> {
    if (this._isFetched) return;

    if (typeof window !== 'undefined') {
      try {
        const spCtx = (window as unknown as { _spPageContextInfo?: { webAbsoluteUrl?: string } })._spPageContextInfo;
        const targetUrl = siteUrl || (spCtx && spCtx.webAbsoluteUrl) || '';

        if (targetUrl) {
          // Attempt 1: v2.1 TermStore Groups with Sets and Terms expanded
          const endpoint = `${targetUrl}/_api/v2.1/termStore/groups?$expand=sets($expand=terms)`;
          const response = await fetch(endpoint, {
            headers: {
              Accept: 'application/json;odata=nometadata',
              'Content-Type': 'application/json'
            },
            credentials: 'include'
          });

          if (response.ok) {
            const data = await response.json();
            if (data && Array.isArray(data.value) && data.value.length > 0) {
              const liveGroups: ITermGroup[] = [];
              data.value.forEach((g: {
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
              }) => {
                const groupName = g.displayName || g.name || 'Global Group';
                let termSets: ITermSet[] = [];

                if (Array.isArray(g.sets) && g.sets.length > 0) {
                  g.sets.forEach((s) => {
                    const setName = s.localizedNames && s.localizedNames[0] ? s.localizedNames[0].name : 'Term Set';
                    const terms: ITermStoreTag[] = [];

                    if (Array.isArray(s.terms) && s.terms.length > 0) {
                      s.terms.forEach((t) => {
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

                        terms.push({
                          id: t.id,
                          label: primaryLabel,
                          termSetId: s.id,
                          termSetName: setName,
                          path: `${groupName} > ${setName} > ${primaryLabel}`,
                          synonyms: synonyms.length > 0 ? synonyms : undefined
                        });
                      });
                    }

                    // If terms are empty in live response for this set, try merging fallback terms
                    if (terms.length === 0) {
                      const fallbackMatchSet = this._findFallbackTermSet(groupName, setName);
                      if (fallbackMatchSet && fallbackMatchSet.terms.length > 0) {
                        terms.push(...fallbackMatchSet.terms);
                      }
                    }

                    termSets.push({
                      id: s.id,
                      name: setName,
                      groupId: g.id,
                      groupName: groupName,
                      terms: terms
                    });
                  });
                }

                // If live group has no term sets returned by REST (e.g. permission or unexpanded), merge fallback term sets
                if (termSets.length === 0) {
                  const fallbackGrp = this._cachedGroups.find(
                    (cg) => cg.name.toLowerCase() === groupName.toLowerCase()
                  );
                  if (fallbackGrp && fallbackGrp.termSets.length > 0) {
                    termSets = fallbackGrp.termSets;
                  }
                }

                liveGroups.push({
                  id: g.id,
                  name: groupName,
                  termSets: termSets
                });
              });

              if (liveGroups.length > 0) {
                // Ensure Intranet and AMCL Terms have their rich term sets preserved
                this._mergeFallbackGroupsIntoLive(liveGroups);
                this._cachedGroups = liveGroups;
              }
            }
          }
        }
      } catch {
        // Fallback gracefully to tenant defaults
      } finally {
        this._isFetched = true;
      }
    }
  }

  /**
   * Helper to find a fallback term set across cached groups.
   */
  private static _findFallbackTermSet(groupName: string, setName: string): ITermSet | undefined {
    const targetSet = setName.toLowerCase();
    for (const grp of this._cachedGroups) {
      const match = grp.termSets.find((s) => s.name.toLowerCase() === targetSet);
      if (match) return match;
    }
    return undefined;
  }

  /**
   * Merges fallback term sets into live groups to guarantee no empty term sets or missing core groups.
   */
  private static _mergeFallbackGroupsIntoLive(liveGroups: ITermGroup[]): void {
    // 1. Check each live group. If its termSets is empty or missing known sets, augment from cached
    this._cachedGroups.forEach((fallbackGrp) => {
      const liveGrp = liveGroups.find(
        (lg) => lg.name.toLowerCase() === fallbackGrp.name.toLowerCase()
      );

      if (liveGrp) {
        if (!liveGrp.termSets || liveGrp.termSets.length === 0) {
          liveGrp.termSets = [...fallbackGrp.termSets];
        } else {
          // If live group is missing key term sets like 'Our business' or 'Our teams', augment them
          fallbackGrp.termSets.forEach((fbSet) => {
            const hasSet = liveGrp.termSets.some(
              (ls) => ls.name.toLowerCase() === fbSet.name.toLowerCase()
            );
            if (!hasSet) {
              liveGrp.termSets.push(fbSet);
            }
          });
        }
      } else {
        // Fallback group wasn't even in live results (e.g. Intranet or Our business)
        liveGroups.push(fallbackGrp);
      }
    });
  }

  /**
   * Returns all hierarchical term groups and their term sets.
   */
  public static async getTermGroups(siteUrl?: string): Promise<ITermGroup[]> {
    await this.initializeFromSharePoint(siteUrl);
    return this._cachedGroups;
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
        s.terms.forEach((t) => {
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
   * Returns list of unique term sets available across all term groups.
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
   * Fetches all terms in a given term set by name (case-insensitive) or set ID.
   */
  public static async getTermsByTermSet(termSetName: string, siteUrl?: string): Promise<ITermStoreTag[]> {
    await this.initializeFromSharePoint(siteUrl);
    const target = termSetName.toLowerCase().trim();
    const matches: ITermStoreTag[] = [];

    this._cachedGroups.forEach((g) => {
      g.termSets.forEach((s) => {
        if (s.name.toLowerCase() === target || s.id.toLowerCase() === target) {
          matches.push(...s.terms);
        }
      });
    });

    return matches;
  }

  /**
   * Adds a new custom term to the active term group/set.
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
