/**
 * @file PropertyPaneTermPickerField.tsx
 * @description Custom SPFx Property Pane Field rendering an intuitive Term Store Explorer picker.
 * Features:
 * - Shows an interactive tag icon button to launch a rich Term Store hierarchical modal.
 * - Once selected, displays a clean summary card with the selected Term name and path.
 * - Provides an 'X' remove button that unlinks the term and restores the tag picker trigger.
 * - Wrapped in FluentProvider with zIndex: 1000000 to prevent clipping in SharePoint edit panels.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import {
  FluentProvider,
  webLightTheme,
  Button,
  Caption1,
  Subtitle2,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Tag,
  Input,
  Skeleton,
  SkeletonItem,
  Portal
} from '@fluentui/react-components';
import {
  TagRegular,
  DismissRegular,
  FolderRegular,
  FolderOpenRegular,
  SearchRegular,
  TagMultipleRegular,
  ChevronDownRegular,
  ChevronRightRegular
} from '@fluentui/react-icons';
import { TaxonomyService, ITermGroup } from '../../fullWidthContainer/services/TaxonomyService';
import { ITermStoreTag } from '../../fullWidthContainer/models/IContainerModels';

export interface IPropertyPaneTermPickerFieldProps {
  label: string;
  selectedTermName?: string;
  siteUrl?: string;
  onSelectTerm: (termName: string, termTag?: ITermStoreTag) => void;
  onRemoveTerm: () => void;
}

const TermPickerControl: React.FC<IPropertyPaneTermPickerFieldProps> = ({
  label,
  selectedTermName,
  siteUrl,
  onSelectTerm,
  onRemoveTerm
}) => {
  const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
  const [termGroups, setTermGroups] = React.useState<ITermGroup[]>([]);
  const [allTerms, setAllTerms] = React.useState<ITermStoreTag[]>([]);
  const [expandedGroups, setExpandedGroups] = React.useState<Record<string, boolean>>({});
  const [expandedSets, setExpandedSets] = React.useState<Record<string, boolean>>({});
  const [selectedTarget, setSelectedTarget] = React.useState<{ name: string; tag?: ITermStoreTag } | null>(null);
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [isLoading, setIsLoading] = React.useState<boolean>(false);

  // Load term groups and terms when dialog opens
  React.useEffect(() => {
    if (isModalOpen) {
      setIsLoading(true);
      if (selectedTermName && selectedTermName !== 'none') {
        setSelectedTarget({ name: selectedTermName });
      } else {
        setSelectedTarget(null);
      }
      Promise.all([
        TaxonomyService.getTermGroups(siteUrl),
        TaxonomyService.getTerms(undefined, undefined, siteUrl)
      ])
        .then(([groups, terms]) => {
          setTermGroups(groups);
          setAllTerms(terms);
          // Expand all groups by default so users see their term sets immediately
          const initialExpandedGroups: Record<string, boolean> = {};
          const initialExpandedSets: Record<string, boolean> = {};
          groups.forEach((g) => {
            initialExpandedGroups[g.id] = true;
            if (Array.isArray(g.termSets)) {
              g.termSets.forEach((s) => {
                if (
                  s.name.toLowerCase() === 'our teams' ||
                  s.name.toLowerCase() === 'our sectors' ||
                  s.name.toLowerCase() === 'our capabilities' ||
                  (selectedTermName && (s.name.toLowerCase() === selectedTermName.toLowerCase() || s.terms.some((t) => t.label.toLowerCase() === selectedTermName.toLowerCase())))
                ) {
                  initialExpandedSets[s.id] = true;
                }
              });
            }
          });
          setExpandedGroups(initialExpandedGroups);
          setExpandedSets(initialExpandedSets);
          setIsLoading(false);
        })
        .catch(() => {
          setIsLoading(false);
        });
    }
  }, [isModalOpen]);

  const toggleGroupExpansion = (grpId: string, e?: React.MouseEvent): void => {
    if (e) {
      e.stopPropagation();
    }
    setExpandedGroups((prev) => {
      const current = prev[grpId] !== undefined ? prev[grpId] : true;
      return {
        ...prev,
        [grpId]: !current
      };
    });
  };

  const toggleSetExpansion = (setId: string, e?: React.MouseEvent): void => {
    if (e) {
      e.stopPropagation();
    }
    setExpandedSets((prev) => {
      const current = prev[setId] !== undefined ? prev[setId] : false;
      return {
        ...prev,
        [setId]: !current
      };
    });
  };

  const handlePickTerm = (name: string, tag?: ITermStoreTag): void => {
    onSelectTerm(name, tag);
    setIsModalOpen(false);
  };

  const handleConfirmSelectedTarget = (): void => {
    if (selectedTarget && selectedTarget.name) {
      onSelectTerm(selectedTarget.name, selectedTarget.tag);
      setIsModalOpen(false);
    }
  };

  // Filtered terms for search mode
  const filteredTerms = React.useMemo(() => {
    if (!searchQuery.trim()) return [];
    const q = searchQuery.toLowerCase().trim();
    return allTerms.filter(
      (t) =>
        t.label.toLowerCase().includes(q) ||
        (t.termSetName && t.termSetName.toLowerCase().includes(q)) ||
        (t.path && t.path.toLowerCase().includes(q)) ||
        (Array.isArray(t.synonyms) && t.synonyms.some((s) => s.toLowerCase().includes(q)))
    );
  }, [allTerms, searchQuery]);

  const hasSelectedTerm = selectedTermName && selectedTermName !== 'none' && selectedTermName.trim().length > 0;

  return (
    <div style={{ marginTop: '8px', marginBottom: '8px' }}>
      <label
        style={{
          display: 'block',
          fontSize: '12px',
          fontWeight: 600,
          marginBottom: '4px',
          color: '#323130'
        }}
      >
        {label}
      </label>

      {hasSelectedTerm ? (
        /* Selected Term Summary Card with Clear (X) Action */
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '6px 10px',
            backgroundColor: '#f3f9fd',
            border: '1px solid #c7e0f4',
            borderRadius: '4px',
            gap: '8px'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' }}>
            <TagMultipleRegular style={{ color: '#0078d4', fontSize: '18px', flexShrink: 0 }} />
            <div style={{ overflow: 'hidden' }}>
              <Subtitle2
                style={{
                  display: 'block',
                  fontSize: '13px',
                  fontWeight: 600,
                  color: '#005a9e',
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis'
                }}
              >
                {selectedTermName}
              </Subtitle2>
              <Caption1 style={{ color: '#605e5c', fontSize: '11px' }}>
                Linked term for synonym resolution
              </Caption1>
            </div>
          </div>

          <Button
            size="small"
            appearance="subtle"
            icon={<DismissRegular style={{ fontSize: '14px', color: '#a4262c' }} />}
            onClick={() => onRemoveTerm()}
            title="Unlink Term"
            aria-label="Remove Term Link"
            style={{ minWidth: '24px', padding: '4px' }}
          />
        </div>
      ) : (
        /* Unlinked State: Tag Icon Explorer Button */
        <Button
          appearance="outline"
          size="medium"
          icon={<TagRegular style={{ fontSize: '16px', color: '#0078d4' }} />}
          onClick={() => setIsModalOpen(true)}
          style={{
            width: '100%',
            justifyContent: 'flex-start',
            gap: '8px',
            borderColor: '#c8c6c4',
            backgroundColor: '#ffffff'
          }}
        >
          <span style={{ fontSize: '12px', color: '#323130' }}>Select Term from Term Store...</span>
        </Button>
      )}

      {/* Hierarchical Term Store Explorer Modal */}
      <Portal>
        <FluentProvider theme={webLightTheme}>
          <Dialog
            open={isModalOpen}
            onOpenChange={(_, data) => setIsModalOpen(data.open)}
          >
            <DialogSurface
              style={{
                maxWidth: '620px',
                width: '92vw',
                zIndex: 1000000,
                backgroundColor: '#FFFFFF !important' as any,
                border: '1px solid #d1d1d1',
                boxShadow: '0 24px 48px rgba(0, 0, 0, 0.28) !important' as any,
                borderRadius: '8px',
                padding: '24px',
                color: '#323130'
              }}
            >
            <DialogTitle
              action={
                <Button
                  appearance="subtle"
                  aria-label="close"
                  icon={<DismissRegular />}
                  onClick={() => setIsModalOpen(false)}
                />
              }
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontFamily: '"Segoe UI", sans-serif' }}>
                <TagRegular style={{ color: '#0078d4', fontSize: '20px' }} />
                <span style={{ fontWeight: 600, fontSize: '18px', color: '#323130' }}>Term Store Explorer</span>
              </div>
            </DialogTitle>

            <DialogBody style={{ fontFamily: '"Segoe UI", sans-serif' }}>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '8px' }}>
                <div style={{ color: '#605e5c', fontSize: '13px', lineHeight: '18px' }}>
                  Browse or search the Term Store. Selecting a term set or parent term will enable automatic synonym resolution for all of its child terms.
                </div>

                {/* Real-time search bar */}
                <Input
                  size="medium"
                  contentBefore={<SearchRegular style={{ color: '#0078d4', fontSize: '16px' }} />}
                  placeholder="Search terms, term sets, or synonyms..."
                  value={searchQuery}
                  onChange={(_, data) => setSearchQuery(data.value)}
                  style={{
                    width: '100%',
                    backgroundColor: '#ffffff',
                    border: '1px solid #8a8886',
                    borderRadius: '4px'
                  }}
                />

                {isLoading ? (
                  <Skeleton style={{ padding: '12px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                    <SkeletonItem style={{ height: '24px', width: '80%' }} />
                    <SkeletonItem style={{ height: '24px', width: '60%' }} />
                    <SkeletonItem style={{ height: '24px', width: '90%' }} />
                  </Skeleton>
                ) : searchQuery.trim() ? (
                  /* Search Results View */
                  <div
                    style={{
                      maxHeight: '320px',
                      overflowY: 'auto',
                      border: '1px solid #d1d1d1',
                      borderRadius: '6px',
                      padding: '8px',
                      backgroundColor: '#ffffff',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '4px'
                    }}
                  >
                    {filteredTerms.map((term) => {
                      const isSelected = selectedTarget?.name === term.label;
                      return (
                        <div
                          key={term.id}
                          onClick={() => setSelectedTarget({ name: term.label, tag: term })}
                          onDoubleClick={() => handlePickTerm(term.label, term)}
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            padding: '8px 12px',
                            borderRadius: '4px',
                            cursor: 'pointer',
                            backgroundColor: isSelected ? '#eff6fc' : '#ffffff',
                            border: isSelected ? '1px solid #0078d4' : '1px solid #edebe9',
                            transition: 'all 0.15s ease'
                          }}
                          onMouseEnter={(e) => {
                            if (!isSelected) e.currentTarget.style.backgroundColor = '#f3f9fd';
                          }}
                          onMouseLeave={(e) => {
                            if (!isSelected) e.currentTarget.style.backgroundColor = '#ffffff';
                          }}
                        >
                          <div>
                            <div style={{ fontWeight: 600, fontSize: '13px', color: isSelected ? '#004578' : '#201f1e' }}>
                              🏷️ {term.label}
                            </div>
                            {term.path && (
                              <div style={{ fontSize: '11px', color: '#605e5c', marginTop: '2px' }}>{term.path}</div>
                            )}
                            {Array.isArray(term.synonyms) && term.synonyms.length > 0 && (
                              <div style={{ fontSize: '11px', color: '#0078d4', marginTop: '2px', fontWeight: 500 }}>
                                Synonyms: {term.synonyms.join(', ')}
                              </div>
                            )}
                          </div>
                          <span
                            style={{
                              fontSize: '11px',
                              padding: '2px 8px',
                              borderRadius: '4px',
                              backgroundColor: isSelected ? '#0078d4' : '#f3f2f1',
                              color: isSelected ? '#ffffff' : '#323130',
                              fontWeight: 600
                            }}
                          >
                            {term.termSetName || 'Term'}
                          </span>
                        </div>
                      );
                    })}

                    {filteredTerms.length === 0 && (
                      <div style={{ padding: '24px', textAlign: 'center', color: '#605e5c', fontSize: '13px' }}>
                        No terms or synonyms matched &ldquo;{searchQuery}&rdquo;.
                      </div>
                    )}
                  </div>
                ) : (
                  /* Hierarchical Tree Browser */
                  <div
                    style={{
                      maxHeight: '340px',
                      overflowY: 'auto',
                      border: '1px solid #d1d1d1',
                      borderRadius: '6px',
                      padding: '8px',
                      backgroundColor: '#ffffff',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '4px'
                    }}
                  >
                    {termGroups.map((grp) => {
                      const isGrpExpanded = expandedGroups[grp.id] !== undefined ? expandedGroups[grp.id] : true;
                      const hasTermSets = Array.isArray(grp.termSets) && grp.termSets.length > 0;
                      return (
                        <div key={grp.id} style={{ marginBottom: '4px' }}>
                          {/* Group Header (Click to open / collapse) */}
                          <div
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'space-between',
                              fontWeight: 600,
                              fontSize: '13px',
                              color: '#323130',
                              padding: '8px 10px',
                              cursor: 'pointer',
                              borderRadius: '4px',
                              backgroundColor: isGrpExpanded ? '#f3f2f1' : '#faf9f8',
                              border: '1px solid #edebe9',
                              transition: 'background-color 0.15s ease'
                            }}
                            onClick={(e) => toggleGroupExpansion(grp.id, e)}
                            onMouseEnter={(e) => {
                              e.currentTarget.style.backgroundColor = '#edebe9';
                            }}
                            onMouseLeave={(e) => {
                              e.currentTarget.style.backgroundColor = isGrpExpanded ? '#f3f2f1' : '#faf9f8';
                            }}
                          >
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                              {isGrpExpanded ? (
                                <ChevronDownRegular style={{ fontSize: '14px', color: '#0078d4' }} />
                              ) : (
                                <ChevronRightRegular style={{ fontSize: '14px', color: '#605e5c' }} />
                              )}
                              {isGrpExpanded ? (
                                <FolderOpenRegular style={{ color: '#0078d4', fontSize: '18px' }} />
                              ) : (
                                <FolderRegular style={{ color: '#605e5c', fontSize: '18px' }} />
                              )}
                              <span style={{ userSelect: 'none' }}>{grp.name}</span>
                            </div>
                            <span style={{ fontSize: '11px', color: '#605e5c', fontWeight: 500 }}>
                              {hasTermSets ? `${grp.termSets.length} set${grp.termSets.length === 1 ? '' : 's'}` : '0 sets'}
                            </span>
                          </div>

                          {/* Term Sets inside Group */}
                          {isGrpExpanded && hasTermSets &&
                            grp.termSets.map((set) => {
                              const isExpanded = expandedSets[set.id] !== undefined ? expandedSets[set.id] : false;
                              const isSetSelected = selectedTarget?.name === set.name;
                              const hasTerms = Array.isArray(set.terms) && set.terms.length > 0;
                              return (
                                <div key={set.id} style={{ marginLeft: '16px', marginTop: '4px' }}>
                                  <div
                                    style={{
                                      display: 'flex',
                                      alignItems: 'center',
                                      justifyContent: 'space-between',
                                      padding: '6px 10px',
                                      borderRadius: '4px',
                                      cursor: 'pointer',
                                      backgroundColor: isSetSelected ? '#eff6fc' : isExpanded ? '#faf9f8' : '#ffffff',
                                      border: isSetSelected ? '1px solid #0078d4' : '1px solid #edebe9',
                                      transition: 'all 0.15s ease'
                                    }}
                                    onClick={(e) => {
                                      toggleSetExpansion(set.id, e);
                                      setSelectedTarget({ name: set.name });
                                    }}
                                    onDoubleClick={() => handlePickTerm(set.name)}
                                    onMouseEnter={(e) => {
                                      if (!isSetSelected) e.currentTarget.style.backgroundColor = '#f3f9fd';
                                    }}
                                    onMouseLeave={(e) => {
                                      if (!isSetSelected) e.currentTarget.style.backgroundColor = isExpanded ? '#faf9f8' : '#ffffff';
                                    }}
                                  >
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                      {hasTerms && (
                                        isExpanded ? (
                                          <ChevronDownRegular style={{ fontSize: '12px', color: '#0078d4' }} />
                                        ) : (
                                          <ChevronRightRegular style={{ fontSize: '12px', color: '#605e5c' }} />
                                        )
                                      )}
                                      {isExpanded ? (
                                        <FolderOpenRegular style={{ color: '#0078d4', fontSize: '16px' }} />
                                      ) : (
                                        <FolderRegular style={{ color: '#605e5c', fontSize: '16px' }} />
                                      )}
                                      <span
                                        style={{
                                          fontWeight: 600,
                                          fontSize: '12px',
                                          color: isSetSelected ? '#004578' : '#201f1e',
                                          userSelect: 'none'
                                        }}
                                      >
                                        {set.name} ({set.terms ? set.terms.length : 0})
                                      </span>
                                    </div>

                                    <span
                                      style={{
                                        fontSize: '11px',
                                        padding: '2px 8px',
                                        borderRadius: '4px',
                                        backgroundColor: isSetSelected ? '#0078d4' : '#edebe9',
                                        color: isSetSelected ? '#ffffff' : '#605e5c',
                                        fontWeight: 600,
                                        userSelect: 'none'
                                      }}
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        setSelectedTarget({ name: set.name });
                                      }}
                                    >
                                      Term Set
                                    </span>
                                  </div>

                                  {/* Terms inside Term Set */}
                                  {isExpanded && (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '3px', marginLeft: '16px', marginTop: '3px' }}>
                                      {hasTerms ? (
                                        set.terms.map((term) => {
                                          const isTermSelected = selectedTarget?.name === term.label;
                                          return (
                                            <div
                                              key={term.id}
                                              style={{
                                                display: 'flex',
                                                alignItems: 'center',
                                                justifyContent: 'space-between',
                                                padding: '6px 10px',
                                                borderRadius: '4px',
                                                cursor: 'pointer',
                                                fontSize: '12px',
                                                backgroundColor: isTermSelected ? '#eff6fc' : '#ffffff',
                                                border: isTermSelected ? '1px solid #0078d4' : '1px solid #f3f2f1',
                                                transition: 'all 0.15s ease'
                                              }}
                                              onMouseEnter={(e) => {
                                                if (!isTermSelected) e.currentTarget.style.backgroundColor = '#f3f9fd';
                                              }}
                                              onMouseLeave={(e) => {
                                                if (!isTermSelected) e.currentTarget.style.backgroundColor = '#ffffff';
                                              }}
                                              onClick={(e) => {
                                                e.stopPropagation();
                                                setSelectedTarget({ name: term.label, tag: term });
                                              }}
                                              onDoubleClick={() => handlePickTerm(term.label, term)}
                                            >
                                              <div>
                                                <span style={{ fontWeight: 600, color: isTermSelected ? '#004578' : '#201f1e' }}>
                                                  🏷️ {term.label}
                                                </span>
                                                {Array.isArray(term.synonyms) && term.synonyms.length > 0 && (
                                                  <span style={{ fontSize: '11px', color: '#605e5c', marginLeft: '8px' }}>
                                                    ({term.synonyms.join(', ')})
                                                  </span>
                                                )}
                                              </div>
                                              <span
                                                style={{
                                                  fontSize: '11px',
                                                  padding: '2px 8px',
                                                  borderRadius: '4px',
                                                  backgroundColor: isTermSelected ? '#0078d4' : '#f3f2f1',
                                                  color: isTermSelected ? '#ffffff' : '#605e5c',
                                                  fontWeight: 500
                                                }}
                                              >
                                                {isTermSelected ? 'Selected' : 'Select'}
                                              </span>
                                            </div>
                                          );
                                        })
                                      ) : (
                                        <div style={{ padding: '6px 10px', fontSize: '11px', color: '#8a8886', fontStyle: 'italic' }}>
                                          No terms in this term set
                                        </div>
                                      )}
                                    </div>
                                  )}
                                </div>
                              );
                            })}

                          {/* Fallback indicator if group has no term sets */}
                          {isGrpExpanded && !hasTermSets && (
                            <div style={{ marginLeft: '24px', marginTop: '4px', padding: '6px 10px', fontSize: '12px', color: '#8a8886', fontStyle: 'italic' }}>
                              No term sets found in this group
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </DialogContent>

              <DialogActions style={{ marginTop: '20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderTop: '1px solid #edebe9', paddingTop: '16px', gap: '16px' }}>
                <div style={{ fontSize: '13px', color: '#605e5c', flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {selectedTarget?.name ? (
                    <span>
                      Selected: <strong style={{ color: '#0078d4' }}>{selectedTarget.name}</strong>
                    </span>
                  ) : (
                    <span>No term or set selected</span>
                  )}
                </div>
                <div style={{ display: 'flex', gap: '12px', flexShrink: 0, alignItems: 'center' }}>
                  <Button
                    appearance="secondary"
                    onClick={() => setIsModalOpen(false)}
                    style={{
                      padding: '6px 20px',
                      minWidth: '100px',
                      height: '36px',
                      whiteSpace: 'nowrap',
                      flexShrink: 0,
                      borderRadius: '4px',
                      border: '1px solid #8a8886',
                      backgroundColor: '#ffffff',
                      color: '#323130',
                      fontWeight: 600,
                      cursor: 'pointer'
                    }}
                  >
                    Cancel
                  </Button>
                  <Button
                    appearance="primary"
                    disabled={!selectedTarget?.name}
                    onClick={handleConfirmSelectedTarget}
                    style={{
                      padding: '6px 24px',
                      minWidth: '170px',
                      height: '36px',
                      whiteSpace: 'nowrap',
                      flexShrink: 0,
                      borderRadius: '4px',
                      backgroundColor: selectedTarget?.name ? '#0078d4' : '#f3f2f1',
                      borderColor: selectedTarget?.name ? '#0078d4' : '#f3f2f1',
                      color: selectedTarget?.name ? '#ffffff' : '#a19f9d',
                      fontWeight: 600,
                      cursor: selectedTarget?.name ? 'pointer' : 'default'
                    }}
                  >
                    Select term parent
                  </Button>
                </div>
              </DialogActions>
            </DialogBody>
        </DialogSurface>
      </Dialog>
    </FluentProvider>
  </Portal>
</div>
  );
};

/**
 * Custom SPFx Property Pane Field Factory
 */
export function createPropertyPaneTermPickerField(
  targetProperty: string,
  props: IPropertyPaneTermPickerFieldProps
): IPropertyPaneField<any> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: targetProperty,
      onRender: (domElement: HTMLElement) => {
        ReactDom.render(
          React.createElement(
            FluentProvider,
            { theme: webLightTheme },
            React.createElement(TermPickerControl, props)
          ),
          domElement
        );
      },
      onDispose: (domElement: HTMLElement) => {
        ReactDom.unmountComponentAtNode(domElement);
      }
    }
  };
}
