/**
 * @file IDynamicFilterTypes.ts
 * @description Type definitions for SPFx Dynamic Data communication and filter payloads.
 */

export interface IDynamicFilterPayload {
  /** Free-text search term entered by user */
  searchQuery: string;
  /** Active term or button filters (e.g. { category: 'Segment A' }) */
  selectedFilters: Record<string, string>;
  /** Automatic pre-filters resolved from logged-in user profile */
  preFilters: Record<string, string>;
  /** Combined whitespace-separated token string for card filtering engines */
  combinedFilterString: string;
  /** ID of the sending web part */
  sourceWebPartId: string;
  /** Epoch timestamp of update */
  timestamp: number;
}

export interface IUserProfilePropertyOption {
  key: string;
  label: string;
  source: 'graph' | 'sharePoint';
}

export interface IProfileTermMapping {
  propertyKey: string;
  termSetName: string; // Name or ID of the linked Term Set (e.g. 'Our teams')
}

export interface ITermWithSynonyms {
  id: string;
  name: string; // Canonical display name e.g. "Knowledge Management"
  termSetName: string;
  synonyms: string[]; // e.g. ["TT Company\\Support Services\\KM"]
}

export const KNOWN_USER_PROFILE_PROPERTIES: IUserProfilePropertyOption[] = [
  { key: 'department', label: 'Department', source: 'graph' },
  { key: 'officeLocation', label: 'Office Location / City', source: 'graph' },
  { key: 'city', label: 'City', source: 'graph' },
  { key: 'state', label: 'State / Region', source: 'graph' },
  { key: 'country', label: 'Country', source: 'graph' },
  { key: 'companyName', label: 'Company Name', source: 'graph' },
  { key: 'jobTitle', label: 'Job Title', source: 'graph' },
  { key: 'userRegion', label: 'User Region', source: 'sharePoint' },
  { key: 'extensionAttribute10', label: 'extensionAttribute10 (Country)', source: 'graph' },
  { key: 'extensionAttribute12', label: 'extensionAttribute12 (Start Date)', source: 'graph' }
];

