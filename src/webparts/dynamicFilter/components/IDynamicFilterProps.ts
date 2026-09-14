/**
 * @file IDynamicFilterProps.ts
 * @description React component props for DynamicFilter component.
 */

import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDynamicFilterPayload } from '../../common/models/IDynamicFilterTypes';

export interface IDynamicFilterProps {
  searchPromptText: string;
  showAppliedFilters: boolean;
  showAutocomplete: boolean;
  enableProfilePreFiltering: boolean;
  selectedPreFilterProperties: string[];
  termMappings?: Record<string, string>; // e.g. { department: 'Our teams' }
  resolvedTermSynonymValues?: Record<string, string>; // e.g. { department: 'Knowledge Management' }
  customGreetingText?: string; // legacy fallback
  enableTimeOfDayGreeting?: boolean; // toggle on/off for Morning/Afternoon
  morningGreetingText?: string;
  includeNameInMorning?: boolean;
  afternoonGreetingText?: string;
  includeNameInAfternoon?: boolean;
  genericGreetingText?: string;
  includeNameInGeneric?: boolean;
  includeFirstNameInGreeting?: boolean; // legacy fallback
  searchAlignment: 'left' | 'center' | 'right';
  backgroundColor?: string;
  spfxTheme?: IReadonlyTheme;
  isDarkTheme?: boolean;
  isEditMode: boolean;
  userProfileDetails: Record<string, any>;
  userProfilePhotoUrl: string;
  userFirstName: string;
  userDisplayName: string;
  isSiteAdmin: boolean;
  onFilterChange: (payload: IDynamicFilterPayload) => void;
  onOpenPropertyPane?: () => void;
}

