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
  SkeletonItem
} from '@fluentui/react-components';
import {
  TagRegular,
  DismissRegular,
  FolderRegular,
  FolderOpenRegular,
  SearchRegular,
  TagMultipleRegular
} from '@fluentui/react-icons';
import { TaxonomyService, ITermGroup } from '../../fullWidthContainer/services/TaxonomyService';
import { ITermStoreTag } from '../../fullWidthContainer/models/IContainerModels';

export interface IPropertyPaneTermPickerFieldProps {
  label: string;
  selectedTermName?: string;
  onSelectTerm: (termName: string, termTag?: ITermStoreTag) => void;
  onRemoveTerm: () => void;
}

const TermPickerControl: React.FC<IPropertyPaneTermPickerFieldProps> = ({
  label,
  selectedTermName,
  onSelectTerm,
  onRemoveTerm
}) => {
  const [isModalOpen, setIsModalOpen] = React.useState<boolean>(false);
  const [termGroups, setTermGroups] = React.useState<ITermGroup[]>([]);
  const [allTerms, setAllTerms] = React.useState<ITermStoreTag[]>([]);
  const [expandedSets, setExpandedSets] = React.useState<Record<string, boolean>>({});
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [isLoading, setIsLoading] = React.useState<boolean>(false);

  // Load term groups and terms when dialog opens
  React.useEffect(() => {
    if (isModalOpen) {
      setIsLoading(true);
      Promise.all([
        TaxonomyService.getTermGroups(),
        TaxonomyService.getTerms()
      ])
        .then(([groups, terms]) => {
          setTermGroups(groups);
          setAllTerms(terms);
          // Auto-expand 'Our business' and 'Our teams'
          const initialExpanded: Record<string, boolean> = {};
          groups.forEach((g) => {
            g.termSets.forEach((s) => {
              if (s.name.toLowerCase() === 'our teams' || s.name.toLowerCase() === 'our sectors') {
                initialExpanded[s.id] = true;
              }
            });
          });
          setExpandedSets(initialExpanded);
          setIsLoading(false);
        })
        .catch(() => {
          setIsLoading(false);
        });
    }
  }, [isModalOpen]);

  const toggleSetExpansion = (setId: string): void => {
    setExpandedSets((prev) => ({
      ...prev,
      [setId]: !prev[setId]
    }));
  };

  const handlePickTerm = (name: string, tag?: ITermStoreTag): void => {
    onSelectTerm(name, tag);
    setIsModalOpen(false);
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
      <Dialog
        open={isModalOpen}
        onOpenChange={(_, data) => setIsModalOpen(data.open)}
      >
        <DialogSurface
          style={{
            maxWidth: '560px',
            width: '90vw',
            zIndex: 1000000,
            backgroundColor: '#FFFFFF',
            border: '1px solid #d1d1d1',
            boxShadow: '0 24px 48px rgba(0, 0, 0, 0.28)',
            borderRadius: '8px',
            padding: '20px'
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
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <TagRegular style={{ color: '#0078d4', fontSize: '20px' }} />
              <span>Term Store Explorer</span>
            </div>
          </DialogTitle>

          <DialogBody>
            <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '8px' }}>
              <Caption1 style={{ color: '#605e5c' }}>
                Browse or search the Term Store. Selecting a term set or parent term will enable automatic synonym resolution for all of its child terms.
              </Caption1>

              {/* Real-time search bar */}
              <Input
                size="medium"
                contentBefore={<SearchRegular style={{ color: '#0078d4' }} />}
                placeholder="Search terms, term sets, or synonyms..."
                value={searchQuery}
                onChange={(_, data) => setSearchQuery(data.value)}
                style={{ width: '100%' }}
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
                    maxHeight: '300px',
                    overflowY: 'auto',
                    border: '1px solid #edebe9',
                    borderRadius: '4px',
                    padding: '8px',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '4px'
                  }}
                >
                  {filteredTerms.map((term) => (
                    <div
                      key={term.id}
                      onClick={() => handlePickTerm(term.label, term)}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        padding: '8px 10px',
                        borderRadius: '4px',
                        cursor: 'pointer',
                        backgroundColor: '#ffffff',
                        border: '1px solid #f3f2f1'
                      }}
                      onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#eff6fc')}
                      onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = '#ffffff')}
                    >
                      <div>
                        <div style={{ fontWeight: 600, fontSize: '13px', color: '#201f1e' }}>
                          🏷️ {term.label}
                        </div>
                        {term.path && (
                          <div style={{ fontSize: '11px', color: '#605e5c' }}>{term.path}</div>
                        )}
                        {Array.isArray(term.synonyms) && term.synonyms.length > 0 && (
                          <div style={{ fontSize: '11px', color: '#0078d4', marginTop: '2px' }}>
                            Synonyms: {term.synonyms.join(', ')}
                          </div>
                        )}
                      </div>
                      <Tag size="small" appearance="outline" shape="rounded">
                        {term.termSetName || 'Term'}
                      </Tag>
                    </div>
                  ))}

                  {filteredTerms.length === 0 && (
                    <div style={{ padding: '16px', textAlign: 'center', color: '#605e5c' }}>
                      No terms or synonyms matched &ldquo;{searchQuery}&rdquo;.
                    </div>
                  )}
                </div>
              ) : (
                /* Hierarchical Tree Browser */
                <div
                  style={{
                    maxHeight: '320px',
                    overflowY: 'auto',
                    border: '1px solid #edebe9',
                    borderRadius: '4px',
                    padding: '8px',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '6px'
                  }}
                >
                  {termGroups.map((grp) => (
                    <div key={grp.id} style={{ marginBottom: '6px' }}>
                      {/* Group Header */}
                      <div
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          gap: '6px',
                          fontWeight: 700,
                          fontSize: '13px',
                          color: '#323130',
                          padding: '4px 6px'
                        }}
                      >
                        <FolderOpenRegular style={{ color: '#0078d4' }} />
                        <span>{grp.name}</span>
                      </div>

                      {/* Term Sets inside Group */}
                      {grp.termSets.map((set) => {
                        const isExpanded = !!expandedSets[set.id];
                        return (
                          <div key={set.id} style={{ marginLeft: '16px', marginTop: '2px' }}>
                            <div
                              style={{
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'space-between',
                                padding: '4px 8px',
                                borderRadius: '4px',
                                cursor: 'pointer',
                                backgroundColor: isExpanded ? '#f3f2f1' : 'transparent'
                              }}
                              onClick={() => toggleSetExpansion(set.id)}
                            >
                              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                {isExpanded ? (
                                  <FolderOpenRegular style={{ color: '#0078d4' }} />
                                ) : (
                                  <FolderRegular style={{ color: '#605e5c' }} />
                                )}
                                <span style={{ fontWeight: 600, fontSize: '12px', color: '#201f1e' }}>
                                  {set.name} ({set.terms.length})
                                </span>
                              </div>

                              <Button
                                size="small"
                                appearance="subtle"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handlePickTerm(set.name);
                                }}
                                style={{ fontSize: '11px', height: '22px', padding: '2px 8px' }}
                              >
                                Select Set
                              </Button>
                            </div>

                            {/* Terms inside Term Set */}
                            {isExpanded && (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '2px', marginLeft: '16px', marginTop: '2px' }}>
                                {set.terms.map((term) => (
                                  <div
                                    key={term.id}
                                    style={{
                                      display: 'flex',
                                      alignItems: 'center',
                                      justifyContent: 'space-between',
                                      padding: '4px 8px',
                                      borderRadius: '4px',
                                      cursor: 'pointer',
                                      fontSize: '12px'
                                    }}
                                    onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = '#eff6fc')}
                                    onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
                                    onClick={() => handlePickTerm(term.label, term)}
                                  >
                                    <div>
                                      <span style={{ fontWeight: 500, color: '#201f1e' }}>🏷️ {term.label}</span>
                                      {Array.isArray(term.synonyms) && term.synonyms.length > 0 && (
                                        <span style={{ fontSize: '11px', color: '#605e5c', marginLeft: '8px' }}>
                                          ({term.synonyms.join(', ')})
                                        </span>
                                      )}
                                    </div>
                                    <Tag size="small" appearance="outline" shape="rounded">
                                      Select
                                    </Tag>
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  ))}
                </div>
              )}
            </DialogContent>

            <DialogActions style={{ marginTop: '12px', display: 'flex', justifyContent: 'flex-end' }}>
              <Button appearance="secondary" onClick={() => setIsModalOpen(false)}>
                Cancel
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
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
