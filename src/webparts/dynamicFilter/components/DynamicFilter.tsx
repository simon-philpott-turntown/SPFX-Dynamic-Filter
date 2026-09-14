/**
 * @file DynamicFilter.tsx
 * @description Pure Fluent UI 2 central Dynamic Filter & Search Controller component.
 * Features:
 * 1. Time-of-day greeting (Good morning/afternoon/evening FirstName)
 * 2. Full-fidelity profile avatar & inspection popover (identical to dashboard screenshots)
 * 3. Central Search bar with autocomplete and prompt text
 * 4. Active filter chips display with clear all
 * 5. Dynamic profile pre-filtering integration
 */

import * as React from 'react';
import {
  makeStyles,
  shorthands,
  tokens,
  FluentProvider,
  Input,
  Button,
  Badge,
  Avatar,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Caption1,
  Subtitle2,
  Body1Strong,
  mergeClasses
} from '@fluentui/react-components';
import {
  SearchRegular,
  DismissRegular,
  DismissCircleRegular,
  ShieldCheckmarkRegular,
  FilterRegular
} from '@fluentui/react-icons';
import { IDynamicFilterProps } from './IDynamicFilterProps';
import { IDynamicFilterPayload } from '../../common/models/IDynamicFilterTypes';
import { getFluent2Theme } from '../../common/utils/themeBridge';
import { TaxonomyService } from '../../fullWidthContainer/services/TaxonomyService';

const useStyles = makeStyles({
  root: {
    width: '100%',
    boxSizing: 'border-box',
    ...shorthands.padding('16px', '24px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
    display: 'flex',
    flexDirection: 'column',
    gap: '12px'
  },
  topBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexWrap: 'wrap',
    gap: '16px'
  },
  greetingAndProfileCol: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    flexShrink: 0
  },
  greetingText: {
    fontSize: '1.15rem',
    fontWeight: 400,
    color: tokens.colorNeutralForeground1,
    lineHeight: '1.25'
  },
  searchAndFilterCol: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    flex: 1,
    minWidth: '280px',
    maxWidth: '680px'
  },
  searchRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    width: '100%'
  },
  searchInput: {
    width: '100%',
    '& input': {
      fontSize: '0.95rem'
    }
  },
  appliedFiltersRow: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: '6px',
    marginTop: '2px'
  },
  clearLink: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    fontSize: '0.75rem',
    fontWeight: 600,
    color: tokens.colorBrandForeground1,
    cursor: 'pointer',
    background: 'none',
    border: 'none',
    padding: '2px 4px',
    textDecoration: 'none',
    ':hover': {
      textDecoration: 'underline'
    }
  },
  popoverSurface: {
    padding: '16px',
    width: '340px',
    maxWidth: '90vw',
    backgroundColor: '#FFFFFF !important' as any,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    boxShadow: '0 24px 48px rgba(0, 0, 0, 0.28) !important' as any,
    ...shorthands.borderRadius(tokens.borderRadiusLarge),
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    zIndex: 1000000
  },
  profileHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px'
  },
  profileTitles: {
    display: 'flex',
    flexDirection: 'column'
  }
});

// ---------------------------------------------------------------------------
// Inline Error Boundary — surfaces render errors to console instead of
// silently triggering the SharePoint "Something went wrong" shell.
// ---------------------------------------------------------------------------
interface IErrorBoundaryState { hasError: boolean; error?: Error }
class DynamicFilterErrorBoundary extends React.Component<{ children: React.ReactNode }, IErrorBoundaryState> {
  public state: IErrorBoundaryState = { hasError: false };
  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { hasError: true, error };
  }
  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('[DynamicFilter] Render error:', error, info.componentStack);
  }
  public render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '12px', color: '#d13438', fontFamily: 'monospace', fontSize: '12px' }}>
          <strong>Dynamic Filter failed to render.</strong> Check the browser console for details.
        </div>
      );
    }
    return this.props.children;
  }
}

export const DynamicFilter: React.FC<IDynamicFilterProps> = (props) => {
  const {
    searchPromptText,
    showAppliedFilters,
    showAutocomplete,
    enableProfilePreFiltering,
    selectedPreFilterProperties,
    termMappings,
    resolvedTermSynonymValues,
    customGreetingText,
    enableTimeOfDayGreeting = true,
    morningGreetingText,
    includeNameInMorning = true,
    afternoonGreetingText,
    includeNameInAfternoon = true,
    genericGreetingText,
    includeNameInGeneric = true,
    includeFirstNameInGreeting = true,
    searchAlignment = 'left',
    spfxTheme,
    isDarkTheme,
    userProfileDetails,
    userProfilePhotoUrl,
    userFirstName,
    userDisplayName,
    isSiteAdmin,
    onFilterChange,
    onOpenPropertyPane
  } = props;

  const styles = useStyles();

  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [committedSearchFilters, setCommittedSearchFilters] = React.useState<string[]>([]);
  const [termStoreTerms, setTermStoreTerms] = React.useState<any[]>([]);
  const [showSuggestions, setShowSuggestions] = React.useState<boolean>(false);
  const [activeExternalFilters, setActiveExternalFilters] = React.useState<Record<string, string>>({});
  const [dismissedPreFilterKeys, setDismissedPreFilterKeys] = React.useState<Record<string, boolean>>({});
  const [isProfileOpen, setIsProfileOpen] = React.useState<boolean>(false);

  // Fetch Term Store terms for canonical titles and synonym-based autocomplete
  React.useEffect(() => {
    let isSubscribed = true;
    TaxonomyService.getTerms(undefined, undefined, props.siteUrl)
      .then((terms) => {
        if (isSubscribed && Array.isArray(terms)) {
          setTermStoreTerms(terms);
        }
      })
      .catch((err) => {
        console.warn('[DynamicFilter] Failed to load term store terms for autocomplete:', err);
      });
    return () => {
      isSubscribed = false;
    };
  }, [props.siteUrl]);

  // Compute theme
  const fluentTheme = React.useMemo(() => {
    return getFluent2Theme(spfxTheme, isDarkTheme);
  }, [spfxTheme, isDarkTheme]);

  // Greeting based on configuration:
  // If time-of-day is on: Morning welcome (<12:00) or Afternoon welcome (>=12:00), each with its own append-name toggle.
  // If time-of-day is off: Generic welcome with its own append-name toggle.
  const greeting = React.useMemo(() => {
    const name = userFirstName && userFirstName.trim().length > 0 ? userFirstName.trim() : 'there';

    if (enableTimeOfDayGreeting) {
      const hour = new Date().getHours();
      const isMorning = hour < 12;

      if (isMorning) {
        const text = (morningGreetingText && morningGreetingText.trim().length > 0)
          ? morningGreetingText.trim()
          : (customGreetingText && customGreetingText.trim().length > 0 ? customGreetingText.trim() : 'Good morning');
        const append = includeNameInMorning !== false;
        return append ? `${text}, ${name}` : text;
      } else {
        const text = (afternoonGreetingText && afternoonGreetingText.trim().length > 0)
          ? afternoonGreetingText.trim()
          : 'Good afternoon';
        const append = includeNameInAfternoon !== false;
        return append ? `${text}, ${name}` : text;
      }
    } else {
      const text = (genericGreetingText && genericGreetingText.trim().length > 0)
        ? genericGreetingText.trim()
        : (customGreetingText && customGreetingText.trim().length > 0 ? customGreetingText.trim() : 'Welcome');
      const append = includeNameInGeneric !== false;
      return append ? `${text}, ${name}` : text;
    }
  }, [
    enableTimeOfDayGreeting,
    morningGreetingText,
    includeNameInMorning,
    afternoonGreetingText,
    includeNameInAfternoon,
    genericGreetingText,
    includeNameInGeneric,
    customGreetingText,
    userFirstName
  ]);

  // Compute automatic pre-filters from user profile properties and resolved synonyms, filtering out dismissed keys
  const resolvedPreFilters = React.useMemo(() => {
    const pre: Record<string, string> = {};
    if (enableProfilePreFiltering && selectedPreFilterProperties && selectedPreFilterProperties.length > 0) {
      selectedPreFilterProperties.forEach((propKey) => {
        const cleanKey = propKey.trim();
        if (cleanKey && !dismissedPreFilterKeys[cleanKey]) {
          // 1. Check if an overriding term synonym value was pre-resolved
          if (resolvedTermSynonymValues && resolvedTermSynonymValues[cleanKey]) {
            const mappedVal = String(resolvedTermSynonymValues[cleanKey]).trim();
            if (mappedVal) {
              pre[cleanKey] = mappedVal;
              return;
            }
          }
          // 2. Fall back to the raw user profile attribute
          if (userProfileDetails && userProfileDetails[cleanKey]) {
            const val = String(userProfileDetails[cleanKey]).trim();
            if (val) {
              pre[cleanKey] = val;
            }
          }
        }
      });
    }
    return pre;
  }, [enableProfilePreFiltering, selectedPreFilterProperties, userProfileDetails, resolvedTermSynonymValues, dismissedPreFilterKeys]);

  // Listen to external card filter apply events from other web parts
  React.useEffect(() => {
    const handleExternalFilter = (e: Event): void => {
      const ce = e as CustomEvent<{ sourceItemId?: string; filterValue?: string }>;
      // CRITICAL: Ignore events dispatched by DynamicFilterWebPart itself to break infinite loops
      if (ce.detail && ce.detail.sourceItemId && ce.detail.sourceItemId !== 'DynamicFilterWebPart') {
        setActiveExternalFilters((prev) => {
          const currentVal = prev[ce.detail!.sourceItemId as string];
          const newVal = ce.detail?.filterValue ? ce.detail.filterValue.trim() : '';

          if (currentVal === newVal) {
            return prev; // No state change, avoid re-renders
          }

          const next = { ...prev };
          if (newVal.length > 0) {
            next[ce.detail!.sourceItemId as string] = newVal;
          } else {
            delete next[ce.detail!.sourceItemId as string];
          }
          return next;
        });
      }
    };

    window.addEventListener('dashboard:card-filter-apply', handleExternalFilter);
    return () => {
      window.removeEventListener('dashboard:card-filter-apply', handleExternalFilter);
    };
  }, []);

  // Track previous combined string to prevent infinite notification loop
  const prevBroadcastRef = React.useRef<string | null>(null);

  // Autocomplete suggestions computed from Term Store canonical terms and synonyms
  // Triggers ONLY when showAutocomplete is enabled AND search query is at least 2 characters
  const autocompleteSuggestions = React.useMemo(() => {
    if (showAutocomplete === false) return [];
    const trimmed = searchQuery.trim().toLowerCase();
    if (trimmed.length < 2) return [];

    interface ISuggestionItem {
      termTitle: string;
      matchedSynonym?: string;
      path?: string;
      termSetName?: string;
    }

    const map = new Map<string, ISuggestionItem>();

    termStoreTerms.forEach((t) => {
      if (!t || !t.label) return;
      const labelLower = t.label.toLowerCase();

      // 1. Exact or partial match on canonical Term Title
      if (labelLower.indexOf(trimmed) !== -1) {
        if (!map.has(t.label)) {
          map.set(t.label, {
            termTitle: t.label,
            termSetName: t.termSetName,
            path: t.path
          });
        }
      }

      // 2. Match against any defined Term Synonyms
      if (Array.isArray(t.synonyms)) {
        t.synonyms.forEach((syn: string) => {
          if (syn && syn.toLowerCase().indexOf(trimmed) !== -1) {
            if (!map.has(t.label)) {
              // Provide the canonical Term title as the suggested autocomplete value
              map.set(t.label, {
                termTitle: t.label,
                matchedSynonym: syn,
                termSetName: t.termSetName,
                path: t.path
              });
            }
          }
        });
      }
    });

    return Array.from(map.values()).slice(0, 8); // Limit to top 8 suggestions
  }, [showAutocomplete, searchQuery, termStoreTerms]);

  // Commit a search string into sticky committed filters (via Enter key or autocomplete click)
  const commitSearchTerm = (term: string): void => {
    const trimmed = term.trim();
    if (!trimmed) return;
    if (committedSearchFilters.indexOf(trimmed) === -1) {
      setCommittedSearchFilters((prev) => [...prev, trimmed]);
    }
    setSearchQuery(''); // Clearing the input text does NOT remove the filter from the list!
    setShowSuggestions(false);
  };

  // Keyboard handler on the search input
  const handleSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      const q = searchQuery.trim();
      if (q) {
        e.preventDefault();
        commitSearchTerm(q);
      }
    } else if (e.key === 'Escape') {
      setShowSuggestions(false);
    }
  };

  // Broadcast combined filter payload whenever committedSearchFilters, searchQuery, preFilters, or activeExternalFilters change
  React.useEffect(() => {
    const tokensList: string[] = [];

    // Include committed sticky search filter chips
    committedSearchFilters.forEach((t) => {
      if (t && tokensList.indexOf(t) === -1) {
        tokensList.push(t);
      }
    });

    // Also include any currently typed search query if present
    if (searchQuery.trim().length > 0 && tokensList.indexOf(searchQuery.trim()) === -1) {
      tokensList.push(searchQuery.trim());
    }

    Object.keys(resolvedPreFilters).forEach((k) => {
      const v = resolvedPreFilters[k];
      if (v && tokensList.indexOf(v) === -1) {
        tokensList.push(v);
      }
    });

    Object.keys(activeExternalFilters).forEach((k) => {
      const v = activeExternalFilters[k];
      if (v && tokensList.indexOf(v) === -1) {
        tokensList.push(v);
      }
    });

    const combinedFilterString = tokensList.join(' ');

    // Break loop if the combined filter string is identical to the last broadcast
    if (prevBroadcastRef.current === combinedFilterString) {
      return;
    }
    prevBroadcastRef.current = combinedFilterString;

    const payload: IDynamicFilterPayload = {
      searchQuery: committedSearchFilters.length > 0 ? committedSearchFilters.join(' ') : searchQuery,
      selectedFilters: activeExternalFilters,
      preFilters: resolvedPreFilters,
      combinedFilterString,
      sourceWebPartId: 'DynamicFilterWebPart',
      timestamp: Date.now()
    };

    onFilterChange(payload);

    // Global DOM Event broadcast for immediate consumer pickup
    window.dispatchEvent(
      new CustomEvent('dashboard:card-filter-apply', {
        detail: {
          sourceItemId: 'DynamicFilterWebPart',
          filterValue: combinedFilterString
        }
      })
    );
  }, [committedSearchFilters, searchQuery, resolvedPreFilters, activeExternalFilters]);

  // Remove individual committed search filter chip
  const handleRemoveCommittedSearchFilter = (index: number): void => {
    setCommittedSearchFilters((prev) => prev.filter((_, idx) => idx !== index));
  };

  // Remove individual uncommitted live search filter
  const handleRemoveSearchFilter = (): void => {
    setSearchQuery('');
  };

  const handleRemovePreFilter = (propKey: string): void => {
    setDismissedPreFilterKeys((prev) => ({
      ...prev,
      [propKey]: true
    }));
  };

  const handleRemoveExternalFilter = (filterId: string): void => {
    setActiveExternalFilters((prev) => {
      const next = { ...prev };
      delete next[filterId];
      return next;
    });
  };

  // Clear all filters: resets committed search chips, active search query, external filters, and marks all pre-filters dismissed
  const handleClearAll = (): void => {
    setCommittedSearchFilters([]);
    setSearchQuery('');
    setActiveExternalFilters({});
    if (selectedPreFilterProperties && selectedPreFilterProperties.length > 0) {
      const allDismissed: Record<string, boolean> = {};
      selectedPreFilterProperties.forEach((k) => {
        const clean = k.trim();
        if (clean) allDismissed[clean] = true;
      });
      setDismissedPreFilterKeys(allDismissed);
    } else {
      setDismissedPreFilterKeys({});
    }
    setShowSuggestions(false);
  };

  const hasAnyActiveFilters =
    committedSearchFilters.length > 0 ||
    searchQuery.trim().length > 0 ||
    Object.keys(activeExternalFilters).length > 0 ||
    Object.keys(resolvedPreFilters).length > 0;

  const searchAlignmentStyle: React.CSSProperties = {
    justifyContent: searchAlignment === 'center' ? 'center' : searchAlignment === 'right' ? 'flex-end' : 'flex-start'
  };

  const rootCustomStyle: React.CSSProperties = React.useMemo(() => {
    if (props.backgroundColor && props.backgroundColor !== 'transparent') {
      return { backgroundColor: props.backgroundColor };
    }
    return { backgroundColor: 'transparent' };
  }, [props.backgroundColor]);

  return (
    <DynamicFilterErrorBoundary>
    <FluentProvider theme={fluentTheme} className={styles.root} style={rootCustomStyle}>
      <div className={styles.topBar} style={searchAlignmentStyle}>
        {/* Left Section: Greeting & Profile Avatar */}
        <div className={styles.greetingAndProfileCol}>
          <Popover
            open={isProfileOpen}
            onOpenChange={(_, data) => setIsProfileOpen(data.open)}
            positioning="below-start"
          >
            <PopoverTrigger disableButtonEnhancement>
              <Button
                appearance="subtle"
                style={{ padding: '2px', minWidth: 'auto', borderRadius: '50%' }}
                title="View user profile summary"
                aria-label="User Profile"
              >
                <Avatar
                  name={userDisplayName || 'User'}
                  image={{ src: userProfilePhotoUrl || undefined }}
                  size={36}
                />
              </Button>
            </PopoverTrigger>

            <PopoverSurface
              style={{
                zIndex: 1000000,
                boxShadow: tokens.shadow28,
                padding: '16px',
                maxWidth: '440px',
                minWidth: '320px',
                backgroundColor: '#FFFFFF',
                border: `1px solid ${tokens.colorNeutralStroke2}`,
                borderRadius: tokens.borderRadiusMedium,
                display: 'flex',
                flexDirection: 'column',
                gap: '10px'
              }}
            >
              {/* Profile Header Summary */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <Avatar
                  size={40}
                  name={userDisplayName || 'User'}
                  image={userProfilePhotoUrl ? { src: userProfilePhotoUrl } : undefined}
                />
                <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Caption1 style={{ fontWeight: 700, fontSize: '1rem', lineHeight: '1.25rem', color: tokens.colorNeutralForeground1 }}>
                      {userDisplayName || 'User'}
                    </Caption1>
                    {isSiteAdmin && (
                      <Badge
                        appearance="filled"
                        color="success"
                        size="medium"
                        icon={<ShieldCheckmarkRegular style={{ fontSize: '14px' }} />}
                        style={{
                          fontSize: '0.75rem',
                          fontWeight: 600,
                          padding: '2px 8px',
                          height: '22px',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '4px'
                        }}
                      >
                        Admin
                      </Badge>
                    )}
                  </div>
                  {userProfileDetails?.jobTitle && (
                    <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: '0.78rem' }}>
                      {userProfileDetails.jobTitle}
                    </Caption1>
                  )}
                  {userProfileDetails?.department && (
                    <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: '0.75rem', fontWeight: 600 }}>
                      {userProfileDetails.department}
                    </Caption1>
                  )}
                  {userProfileDetails?.officeLocation && (
                    <Caption1 style={{ color: tokens.colorNeutralForeground4, fontSize: '0.72rem' }}>
                      {userProfileDetails.officeLocation}
                    </Caption1>
                  )}
                </div>
              </div>
            </PopoverSurface>
          </Popover>

          {/* Greeting text to the right of the profile picture (Normal unbolded text, no subtext) */}
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <span className={styles.greetingText}>{greeting}</span>
          </div>
        </div>

        {/* Centre / Right Section: Search bar & Filter controls */}
        <div className={styles.searchAndFilterCol} style={{ position: 'relative' }}>
          <div className={styles.searchRow}>
            <Input
              className={styles.searchInput}
              contentBefore={<SearchRegular style={{ fontSize: '18px', color: tokens.colorBrandForeground1 }} />}
              contentAfter={
                searchQuery ? (
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<DismissCircleRegular />}
                    onClick={() => {
                      setSearchQuery('');
                      setShowSuggestions(false);
                    }}
                    aria-label="Clear search text"
                  />
                ) : undefined
              }
              placeholder={searchPromptText || 'Search dashboard cards, tags, metrics...'}
              value={searchQuery}
              onChange={(_, data) => {
                setSearchQuery(data.value);
                setShowSuggestions(data.value.trim().length >= 2);
              }}
              onFocus={() => {
                if (searchQuery.trim().length >= 2) {
                  setShowSuggestions(true);
                }
              }}
              onKeyDown={handleSearchKeyDown}
              size="medium"
            />
          </div>

          {/* Autocomplete Suggestions Popover Dropdown */}
          {showSuggestions && autocompleteSuggestions.length > 0 && (
            <div
              style={{
                position: 'absolute',
                top: '42px',
                left: 0,
                right: 0,
                zIndex: 1000000,
                backgroundColor: '#FFFFFF',
                borderRadius: tokens.borderRadiusMedium,
                boxShadow: tokens.shadow16,
                border: `1px solid ${tokens.colorNeutralStroke1}`,
                padding: '4px',
                display: 'flex',
                flexDirection: 'column',
                gap: '2px',
                maxHeight: '260px',
                overflowY: 'auto'
              }}
            >
              {autocompleteSuggestions.map((item, idx) => (
                <div
                  key={`${item.termTitle}-${idx}`}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '8px 12px',
                    borderRadius: '4px',
                    cursor: 'pointer',
                    fontSize: '13px',
                    color: tokens.colorNeutralForeground1,
                    transition: 'background-color 0.15s ease'
                  }}
                  onMouseDown={(e) => {
                    e.preventDefault(); // Prevent blurring input
                    commitSearchTerm(item.termTitle);
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = '#eff6fc';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = 'transparent';
                  }}
                >
                  <div style={{ display: 'flex', flexDirection: 'column' }}>
                    <div style={{ fontWeight: 600, color: '#004578' }}>
                      {item.termTitle}
                    </div>
                    {item.matchedSynonym && (
                      <span style={{ fontSize: '11px', color: '#605e5c', fontStyle: 'italic' }}>
                        Matches synonym: &ldquo;{item.matchedSynonym}&rdquo;
                      </span>
                    )}
                  </div>
                  {item.termSetName && (
                    <span
                      style={{
                        fontSize: '11px',
                        padding: '2px 6px',
                        backgroundColor: '#f3f2f1',
                        borderRadius: '4px',
                        color: '#605e5c'
                      }}
                    >
                      {item.termSetName}
                    </span>
                  )}
                </div>
              ))}
            </div>
          )}

          {/* Active Applied Filters Chips Row */}
          {showAppliedFilters && hasAnyActiveFilters && (
            <div className={styles.appliedFiltersRow}>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, fontWeight: 600 }}>
                Filters applied:
              </Caption1>

              {/* Committed Sticky Search Filter Chips (Sticks upon pressing Enter) */}
              {committedSearchFilters.map((filterText, idx) => (
                <span
                  key={`committed-${filterText}-${idx}`}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px',
                    padding: '2px 8px',
                    borderRadius: '12px',
                    backgroundColor: tokens.colorBrandBackground2,
                    color: tokens.colorBrandForeground2,
                    fontSize: '0.75rem',
                    fontWeight: 500,
                    border: `1px solid ${tokens.colorBrandStroke2}`
                  }}
                >
                  <SearchRegular style={{ fontSize: '12px' }} />
                  <span>{filterText}</span>
                  <DismissRegular
                    onClick={() => handleRemoveCommittedSearchFilter(idx)}
                    role="button"
                    tabIndex={0}
                    aria-label={`Remove filter ${filterText}`}
                    title={`Remove ${filterText}`}
                    style={{
                      cursor: 'pointer',
                      fontSize: '11px',
                      marginLeft: '2px',
                      opacity: 0.8
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        handleRemoveCommittedSearchFilter(idx);
                      }
                    }}
                  />
                </span>
              ))}

              {/* Live Search Query Chip if user has typed something not yet committed with Enter */}
              {searchQuery.trim() && committedSearchFilters.indexOf(searchQuery.trim()) === -1 && (
                <span
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px',
                    padding: '2px 8px',
                    borderRadius: '12px',
                    backgroundColor: tokens.colorBrandBackground2,
                    color: tokens.colorBrandForeground2,
                    fontSize: '0.75rem',
                    fontWeight: 500,
                    border: `1px dashed ${tokens.colorBrandStroke2}`
                  }}
                  title="Press Enter to commit filter"
                >
                  <SearchRegular style={{ fontSize: '12px' }} />
                  <span>{searchQuery.trim()}</span>
                  <DismissRegular
                    onClick={handleRemoveSearchFilter}
                    role="button"
                    tabIndex={0}
                    aria-label="Remove active search text"
                    title="Remove active search text"
                    style={{
                      cursor: 'pointer',
                      fontSize: '11px',
                      marginLeft: '2px',
                      opacity: 0.8
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        handleRemoveSearchFilter();
                      }
                    }}
                  />
                </span>
              )}

              {/* Pre-filters */}
              {Object.keys(resolvedPreFilters).map((propKey) => (
                <span
                  key={propKey}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px',
                    padding: '2px 8px',
                    borderRadius: '12px',
                    backgroundColor: tokens.colorNeutralBackground3,
                    color: tokens.colorNeutralForeground1,
                    fontSize: '0.75rem',
                    fontWeight: 500,
                    border: `1px solid ${tokens.colorNeutralStroke2}`
                  }}
                >
                  <span>{resolvedPreFilters[propKey]}</span>
                  <DismissRegular
                    onClick={() => handleRemovePreFilter(propKey)}
                    role="button"
                    tabIndex={0}
                    aria-label={`Remove filter for ${propKey}`}
                    title={`Remove ${resolvedPreFilters[propKey]}`}
                    style={{
                      cursor: 'pointer',
                      fontSize: '11px',
                      marginLeft: '2px',
                      opacity: 0.8
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        handleRemovePreFilter(propKey);
                      }
                    }}
                  />
                </span>
              ))}

              {/* External Filters */}
              {Object.keys(activeExternalFilters).map((filterId) => (
                <span
                  key={filterId}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px',
                    padding: '2px 8px',
                    borderRadius: '12px',
                    backgroundColor: tokens.colorNeutralBackground3,
                    color: tokens.colorNeutralForeground2,
                    fontSize: '0.75rem',
                    fontWeight: 500,
                    border: `1px solid ${tokens.colorNeutralStroke2}`
                  }}
                >
                  <span>{activeExternalFilters[filterId]}</span>
                  <DismissRegular
                    onClick={() => handleRemoveExternalFilter(filterId)}
                    role="button"
                    tabIndex={0}
                    aria-label={`Remove filter ${filterId}`}
                    title={`Remove ${activeExternalFilters[filterId]}`}
                    style={{
                      cursor: 'pointer',
                      fontSize: '11px',
                      marginLeft: '2px',
                      opacity: 0.8
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        handleRemoveExternalFilter(filterId);
                      }
                    }}
                  />
                </span>
              ))}

              <button
                type="button"
                className={styles.clearLink}
                onClick={handleClearAll}
                title="Clear all active filters"
              >
                <DismissRegular style={{ fontSize: '12px' }} />
                Clear all
              </button>
            </div>
          )}
        </div>
      </div>
    </FluentProvider>
    </DynamicFilterErrorBoundary>
  );
};


