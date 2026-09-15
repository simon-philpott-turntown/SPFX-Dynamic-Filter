/**
 * @file DynamicFilterWebPart.ts
 * @description Central SPFx Dynamic Data Web Part capturing search & filter values,
 * broadcasting them to connected dashboard web parts, showing a time-of-day greeting,
 * and performing user profile pre-filtering. Updated: 2026-09-15.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneFieldType,
  type IPropertyPaneField,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

import { DynamicFilter } from './components/DynamicFilter';
import { IDynamicFilterProps } from './components/IDynamicFilterProps';
import { IDynamicFilterPayload, KNOWN_USER_PROFILE_PROPERTIES } from '../common/models/IDynamicFilterTypes';
import { UserProfileHarvesterService, IUserProfileHarvestResult } from '../common/services/UserProfileHarvesterService';
import { BrandColorPickerPopover } from '../fullWidthContainer/components/BrandColorPickerPopover';
import { TaxonomyService } from '../fullWidthContainer/services/TaxonomyService';
import { createPropertyPaneTermPickerField } from './components/PropertyPaneTermPickerField';

export interface IDynamicFilterWebPartProps {
  searchPromptText: string;
  showAppliedFilters: boolean;
  showAutocomplete: boolean;
  enableProfilePreFiltering: boolean;
  selectedPreFilterProperties: string;
  searchAlignment: 'left' | 'center' | 'right';
  backgroundColor?: string;
  // Greeting properties
  enableTimeOfDayGreeting?: boolean;
  morningGreetingText?: string;
  includeNameInMorning?: boolean;
  afternoonGreetingText?: string;
  includeNameInAfternoon?: boolean;
  genericGreetingText?: string;
  includeNameInGeneric?: boolean;
  customGreetingText?: string; // legacy fallback
  includeFirstNameInGreeting?: boolean; // legacy fallback
  // Dynamic profile pre-filtering and Term Store linking
  termMappingsJson?: string;
  [key: string]: any; // Allow dynamic property keys such as prefilter_department, termlink_department
}

/**
 * Creates an SPFx custom property pane field rendering the visual BrandColorPickerPopover.
 */
const createColorPickerPropertyField = (
  key: string,
  selectedColor: string | undefined,
  onChange: (color?: string) => void,
  defaultLabel: string,
  defaultColorHex: string = 'transparent'
): IPropertyPaneField<any> => ({
  type: PropertyPaneFieldType.Custom,
  targetProperty: key,
  properties: {
    key,
    onRender: (domElement: HTMLElement) => {
      ReactDom.render(
        React.createElement(
          FluentProvider,
          { theme: webLightTheme },
          React.createElement(BrandColorPickerPopover, {
            selectedColor,
            onChange,
            defaultLabel,
            defaultColorHex
          })
        ),
        domElement
      );
    },
    onDispose: (domElement: HTMLElement) => {
      ReactDom.unmountComponentAtNode(domElement);
    }
  }
});

/**
 * Custom Property Pane Field rendering a checkbox with a primary bold label and a second line in smaller grey italics.
 */
const createTwoLineCheckboxPropertyField = (
  key: string,
  title: string,
  subtext: string,
  checked: boolean,
  onChange: (newChecked: boolean) => void
): IPropertyPaneField<any> => ({
  type: PropertyPaneFieldType.Custom,
  targetProperty: key,
  properties: {
    key,
    onRender: (domElement: HTMLElement) => {
      domElement.innerHTML = '';
      const container = document.createElement('div');
      container.style.display = 'flex';
      container.style.alignItems = 'flex-start';
      container.style.gap = '8px';
      container.style.margin = '6px 0';
      container.style.cursor = 'pointer';

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.id = `input_${key}`;
      checkbox.checked = checked;
      checkbox.style.marginTop = '3px';
      checkbox.style.cursor = 'pointer';

      const textContainer = document.createElement('label');
      textContainer.htmlFor = `input_${key}`;
      textContainer.style.display = 'flex';
      textContainer.style.flexDirection = 'column';
      textContainer.style.gap = '2px';
      textContainer.style.cursor = 'pointer';
      textContainer.style.userSelect = 'none';

      const titleSpan = document.createElement('span');
      titleSpan.style.fontSize = '13px';
      titleSpan.style.fontWeight = '600';
      titleSpan.style.color = '#323130';
      titleSpan.innerText = title;

      const subtextSpan = document.createElement('span');
      subtextSpan.style.fontSize = '11px';
      subtextSpan.style.fontStyle = 'italic';
      subtextSpan.style.color = '#605e5c';
      subtextSpan.innerText = subtext;

      textContainer.appendChild(titleSpan);
      textContainer.appendChild(subtextSpan);

      checkbox.onchange = () => {
        onChange(checkbox.checked);
      };

      container.appendChild(checkbox);
      container.appendChild(textContainer);
      domElement.appendChild(container);
    },
    onDispose: (domElement: HTMLElement) => {
      domElement.innerHTML = '';
    }
  }
});

export default class DynamicFilterWebPart
  extends BaseClientSideWebPart<IDynamicFilterWebPartProps>
  implements IDynamicDataCallables {

  private _isDarkTheme: boolean = false;
  private _themeVariant: IReadonlyTheme | undefined;
  private _userHarvestResult: IUserProfileHarvestResult | undefined;
  private _resolvedTermSynonyms: Record<string, string> = {};
  private _currentPayload: IDynamicFilterPayload = {
    searchQuery: '',
    selectedFilters: {},
    preFilters: {},
    combinedFilterString: '',
    sourceWebPartId: 'DynamicFilterWebPart',
    timestamp: Date.now()
  };

  public async onInit(): Promise<void> {
    await super.onInit();

    // Defensive fallback: ensure this.properties is populated with manifest defaults
    if (!this.properties) {
      (this as any).properties = {};
    }
    if (this.properties.selectedPreFilterProperties === undefined) {
      this.properties.selectedPreFilterProperties = 'department,officeLocation';
    }
    if (this.properties.searchPromptText === undefined) {
      this.properties.searchPromptText = 'Search all dashboard cards, tags, metrics...';
    }
    if (this.properties.showAppliedFilters === undefined) {
      this.properties.showAppliedFilters = true;
    }
    if (this.properties.showAutocomplete === undefined) {
      this.properties.showAutocomplete = true;
    }
    if (this.properties.enableProfilePreFiltering === undefined) {
      this.properties.enableProfilePreFiltering = true;
    }
    if (this.properties.searchAlignment === undefined) {
      this.properties.searchAlignment = 'left';
    }
    if (this.properties.backgroundColor === undefined) {
      this.properties.backgroundColor = 'transparent';
    }
    if (this.properties.enableTimeOfDayGreeting === undefined) {
      this.properties.enableTimeOfDayGreeting = true;
    }
    if (this.properties.morningGreetingText === undefined) {
      this.properties.morningGreetingText = 'Good morning';
    }
    if (this.properties.includeNameInMorning === undefined) {
      this.properties.includeNameInMorning = true;
    }
    if (this.properties.afternoonGreetingText === undefined) {
      this.properties.afternoonGreetingText = 'Good afternoon';
    }
    if (this.properties.includeNameInAfternoon === undefined) {
      this.properties.includeNameInAfternoon = true;
    }
    if (this.properties.genericGreetingText === undefined) {
      this.properties.genericGreetingText = 'Welcome';
    }
    if (this.properties.includeNameInGeneric === undefined) {
      this.properties.includeNameInGeneric = true;
    }
    if (this.properties.customGreetingText === undefined) {
      this.properties.customGreetingText = 'Good morning';
    }
    if (this.properties.includeFirstNameInGreeting === undefined) {
      this.properties.includeFirstNameInGreeting = true;
    }

    // Initialize individual checkbox properties from selectedPreFilterProperties if not already set
    const activePropsList = (this.properties.selectedPreFilterProperties || '')
      .split(',')
      .map((s) => s.trim().toLowerCase())
      .filter((s) => s.length > 0);

    KNOWN_USER_PROFILE_PROPERTIES.forEach((prop) => {
      const checkboxKey = `prefilter_${prop.key}`;
      if (this.properties[checkboxKey] === undefined) {
        this.properties[checkboxKey] = activePropsList.indexOf(prop.key.toLowerCase()) !== -1;
      }
    });

    // Default term mapping for department -> 'Our teams' as per the user's workflow
    if (this.properties['termlink_department'] === undefined) {
      this.properties['termlink_department'] = 'Our teams';
    }

    // Register this web part as an SPFx Dynamic Data Source.
    try {
      this.context.dynamicDataSourceManager.initializeSource(this);
    } catch (initErr) {
      console.warn('[DynamicFilterWebPart] Dynamic data source registration failed (non-fatal):', initErr);
    }

    // Pre-seed user details synchronously from cached profile or pageContext for instant 0ms greeting, avatar, and pre-filters
    const cachedProfile = UserProfileHarvesterService.getCachedResult(this.context);
    const currentUser = this.context?.pageContext?.user;
    const legacyCtx = (window as any)._spPageContextInfo || (this.context?.pageContext as any)?.legacyPageContext;
    const isSiteAdmin = Boolean(legacyCtx?.isSiteAdmin);
    const userAccount = currentUser?.email || currentUser?.loginName || '';
    const webRelativeUrl = this.context?.pageContext?.web?.serverRelativeUrl;
    const normalizedWebUrl = webRelativeUrl === '/' ? '' : (webRelativeUrl || '');
    const instantPhotoUrl = cachedProfile?.photoUrl || (userAccount
      ? `${normalizedWebUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(userAccount)}`
      : '');
    const instantFirstName = cachedProfile?.firstName || (currentUser?.displayName ? currentUser.displayName.split(' ')[0] : 'there');

    this._userHarvestResult = cachedProfile || {
      details: {
        DisplayName: currentUser?.displayName || '',
        email: currentUser?.email || '',
        loginName: currentUser?.loginName || '',
        FirstName: instantFirstName,
        isSiteAdmin
      },
      photoUrl: instantPhotoUrl,
      firstName: instantFirstName,
      displayName: currentUser?.displayName || 'Colleague',
      isSiteAdmin
    };

    // Pre-seed taxonomy service and resolve synonyms as soon as live taxonomy is ready
    void TaxonomyService.initializeFromSharePoint(this.context?.pageContext?.web?.absoluteUrl)
      .then(async () => {
        await this._resolveAllSynonyms();
        this.render();
      })
      .catch((taxErr) => {
        console.warn('[DynamicFilterWebPart] Taxonomy pre-seed error (non-fatal):', taxErr);
      });

    // Harvest user profile details & avatar photo in background (non-blocking).
    void UserProfileHarvesterService.harvest(this.context)
      .then(async (result) => {
        this._userHarvestResult = result;
        await this._resolveAllSynonyms();
        this.render();
        if (this.context?.propertyPane?.isPropertyPaneOpen()) {
          this.context.propertyPane.refresh();
        }
      })
      .catch((err) => {
        console.warn('[DynamicFilterWebPart] Background harvest error (non-fatal):', err);
      });
  }

  /**
   * Resolves any linked Term Set synonyms against current user profile values.
   */
  private async _resolveAllSynonyms(): Promise<void> {
    if (!this._userHarvestResult?.details) return;
    const details = this._userHarvestResult.details;
    const resolved: Record<string, string> = {};

    for (const prop of KNOWN_USER_PROFILE_PROPERTIES) {
      const isEnabled = this.properties[`prefilter_${prop.key}`] !== false;
      const linkedTermSet = this.properties[`termlink_${prop.key}`];
      const rawUserVal = details[prop.key];

      if (isEnabled && linkedTermSet && linkedTermSet !== 'none' && rawUserVal) {
        try {
          const resolvedVal = await TaxonomyService.resolveTermFromSynonym(
            String(rawUserVal),
            linkedTermSet,
            this.context?.pageContext?.web?.absoluteUrl
          );
          if (resolvedVal && resolvedVal !== String(rawUserVal)) {
            resolved[prop.key] = resolvedVal;
          }
        } catch (synonymErr) {
          console.warn(`[DynamicFilterWebPart] Failed to resolve synonym for ${prop.key}:`, synonymErr);
        }
      }
    }
    this._resolvedTermSynonyms = resolved;
  }

  /**
   * Return the list of properties that this web part can share with consumers.
   */
  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'filterPayload',
        title: 'Filter Payload (Full Object)'
      },
      {
        id: 'searchQuery',
        title: 'Search Query Text'
      },
      {
        id: 'combinedFilterString',
        title: 'Combined Filter Tokens'
      }
    ];
  }

  /**
   * Returns value for the requested dynamic data property.
   */
  public getPropertyValue(propertyId: string): any {
    switch (propertyId) {
      case 'filterPayload':
        return this._currentPayload;
      case 'searchQuery':
        return this._currentPayload.searchQuery;
      case 'combinedFilterString':
        return this._currentPayload.combinedFilterString;
      default:
        return undefined;
    }
  }

  public render(): void {
    if (!this.domElement) {
      return;
    }

    try {
      if (!this.properties) {
        (this as any).properties = {};
      }
      const props = this.properties;

      // Compile active properties from individual checkbox toggles
      const activeProperties: string[] = [];
      KNOWN_USER_PROFILE_PROPERTIES.forEach((prop) => {
        if (props[`prefilter_${prop.key}`] === true) {
          activeProperties.push(prop.key);
        }
      });

      // Keep selectedPreFilterProperties comma-string synchronized for backwards compatibility
      if (this.properties) {
        this.properties.selectedPreFilterProperties = activeProperties.join(',');
      }

      // Build term mappings dictionary
      const termMappings: Record<string, string> = {};
      KNOWN_USER_PROFILE_PROPERTIES.forEach((prop) => {
        const linkedSet = props[`termlink_${prop.key}`];
        if (linkedSet && linkedSet !== 'none') {
          termMappings[prop.key] = linkedSet;
        }
      });

      const element: React.ReactElement<IDynamicFilterProps> = React.createElement(DynamicFilter, {
        searchPromptText: props.searchPromptText || 'Search all dashboard cards, tags, metrics...',
        showAppliedFilters: props.showAppliedFilters !== false,
        showAutocomplete: props.showAutocomplete !== false,
        enableProfilePreFiltering: props.enableProfilePreFiltering !== false,
        selectedPreFilterProperties: activeProperties,
        termMappings,
        resolvedTermSynonymValues: this._resolvedTermSynonyms,
        enableTimeOfDayGreeting: props.enableTimeOfDayGreeting !== false,
        morningGreetingText: props.morningGreetingText || 'Good morning',
        includeNameInMorning: props.includeNameInMorning !== false,
        afternoonGreetingText: props.afternoonGreetingText || 'Good afternoon',
        includeNameInAfternoon: props.includeNameInAfternoon !== false,
        genericGreetingText: props.genericGreetingText || 'Welcome',
        includeNameInGeneric: props.includeNameInGeneric !== false,
        customGreetingText: props.customGreetingText || 'Good morning',
        includeFirstNameInGreeting: props.includeFirstNameInGreeting !== false,
        searchAlignment: props.searchAlignment || 'left',
        backgroundColor: props.backgroundColor,
        spfxTheme: this._themeVariant,
        isDarkTheme: this._isDarkTheme,
        isEditMode: this.displayMode === 2, // 2 = Edit Mode
        userProfileDetails: this._userHarvestResult?.details || {},
        userProfilePhotoUrl: this._userHarvestResult?.photoUrl || '',
        userFirstName: this._userHarvestResult?.firstName || '',
        userDisplayName: this._userHarvestResult?.displayName || '',
        isSiteAdmin: this._userHarvestResult?.isSiteAdmin || false,
        siteUrl: this.context?.pageContext?.web?.absoluteUrl,
        onFilterChange: (payload: IDynamicFilterPayload) => {
          this._currentPayload = payload;
          // Notify SPFx Dynamic Data consumers
          try {
            this.context.dynamicDataSourceManager.notifyPropertyChanged('filterPayload');
            this.context.dynamicDataSourceManager.notifyPropertyChanged('searchQuery');
            this.context.dynamicDataSourceManager.notifyPropertyChanged('combinedFilterString');
          } catch (notifyErr) {
            console.warn('[DynamicFilterWebPart] notifyPropertyChanged error (non-fatal):', notifyErr);
          }
        },
        onOpenPropertyPane: () => this.context.propertyPane.open()
      });

      ReactDom.render(element, this.domElement);
    } catch (err) {
      console.error('[DynamicFilterWebPart] Critical Render Error:', err);
      const errorDetails = err instanceof Error ? (err.stack || err.message) : (typeof err === 'object' ? JSON.stringify(err, Object.getOwnPropertyNames(err), 2) : String(err));
      this.domElement.innerHTML = `
        <div style="padding: 24px; color: #a80000; background: #fde7e9; border: 2px solid #d13438; border-radius: 8px; font-family: Segoe UI, sans-serif;">
          <h3 style="margin-top:0; font-size: 18px;">⚠️ Dynamic Filter Web Part Render Error</h3>
          <p style="font-size: 14px; margin-bottom: 12px;">An exception occurred inside the web part render cycle:</p>
          <pre style="white-space: pre-wrap; word-break: break-all; background: #ffffff; padding: 12px; border: 1px solid #d13438; border-radius: 4px; font-size: 12px; color: #333333;">${errorDetails}</pre>
        </div>
      `;
    }
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'enableTimeOfDayGreeting') {
      this.render();
      if (this.context?.propertyPane?.isPropertyPaneOpen()) {
        this.context.propertyPane.refresh();
      }
    } else if (propertyPath.startsWith('termlink_') || propertyPath.startsWith('prefilter_')) {
      await this._resolveAllSynonyms();
      this.render();
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._themeVariant = currentTheme;
    this._isDarkTheme = !!currentTheme.isInverted;
    if (this.domElement && this.properties) {
      this.render();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const props = this.properties || ({} as Partial<IDynamicFilterWebPartProps>);
    const userDetails = this._userHarvestResult?.details || {};
    const isTimeOfDayOn = props.enableTimeOfDayGreeting !== false;

    // Build greeting configuration fields dynamically based on enableTimeOfDayGreeting toggle
    const greetingFields: IPropertyPaneField<any>[] = [
      PropertyPaneToggle('enableTimeOfDayGreeting', {
        label: 'Time-of-day salutation (Morning / Afternoon / Evening)',
        checked: isTimeOfDayOn
      })
    ];

    if (isTimeOfDayOn) {
      // Morning welcome group
      greetingFields.push(
        PropertyPaneLabel('morningWelcomeHeading', {
          text: 'Morning welcome'
        }),
        PropertyPaneTextField('morningGreetingText', {
          label: 'Morning message',
          value: props.morningGreetingText || 'Good morning'
        }),
        PropertyPaneCheckbox('includeNameInMorning', {
          text: 'Append name to morning message',
          checked: props.includeNameInMorning !== false
        }),
        PropertyPaneHorizontalRule(),
        // Afternoon welcome group
        PropertyPaneLabel('afternoonWelcomeHeading', {
          text: 'Afternoon welcome'
        }),
        PropertyPaneTextField('afternoonGreetingText', {
          label: 'Afternoon message',
          value: props.afternoonGreetingText || 'Good afternoon'
        }),
        PropertyPaneCheckbox('includeNameInAfternoon', {
          text: 'Append name to afternoon message',
          checked: props.includeNameInAfternoon !== false
        })
      );
    } else {
      // Generic welcome group
      greetingFields.push(
        PropertyPaneLabel('genericWelcomeHeading', {
          text: 'Generic welcome'
        }),
        PropertyPaneTextField('genericGreetingText', {
          label: 'Generic message',
          value: props.genericGreetingText || 'Welcome'
        }),
        PropertyPaneCheckbox('includeNameInGeneric', {
          text: 'Append name to generic message',
          checked: props.includeNameInGeneric !== false
        })
      );
    }

    // Generate dynamic property fields for each user profile attribute with 2-line formatting
    const dynamicPreFilterFields: IPropertyPaneField<any>[] = [
      PropertyPaneToggle('enableProfilePreFiltering', {
        label: 'Enable user profile auto-filtering on page load',
        checked: props.enableProfilePreFiltering !== false
      }),
      PropertyPaneLabel('preFilterInfo', {
        text: 'Tick the checkbox for each property to use it as an active filter. Click the tag icon to select a Term or Term Set: any match in term synonyms will automatically resolve to the canonical Term Name.'
      }),
      PropertyPaneHorizontalRule()
    ];

    KNOWN_USER_PROFILE_PROPERTIES.forEach((prop) => {
      const rawUserVal = userDetails[prop.key];
      const hasValue = rawUserVal !== undefined && rawUserVal !== null && String(rawUserVal).trim().length > 0;
      const subtext = hasValue ? `(e.g. ${String(rawUserVal)})` : '(e.g. not set)';
      const resolvedSynonym = this._resolvedTermSynonyms[prop.key];
      const resolvedNote = resolvedSynonym ? ` ➜ Resolves to Term: "${resolvedSynonym}"` : '';

      dynamicPreFilterFields.push(
        createTwoLineCheckboxPropertyField(
          `prefilter_${prop.key}`,
          prop.label,
          subtext,
          props[`prefilter_${prop.key}`] === true,
          (newChecked: boolean) => {
            this.properties[`prefilter_${prop.key}`] = newChecked;
            this.render();
          }
        ),
        createPropertyPaneTermPickerField(`termlink_${prop.key}`, {
          label: `Link ${prop.label} to Term Store`,
          selectedTermName: props[`termlink_${prop.key}`],
          siteUrl: this.context?.pageContext?.web?.absoluteUrl,
          onSelectTerm: async (termName: string) => {
            this.properties[`termlink_${prop.key}`] = termName;
            await this._resolveAllSynonyms();
            this.render();
            if (this.context?.propertyPane?.isPropertyPaneOpen()) {
              this.context.propertyPane.refresh();
            }
          },
          onRemoveTerm: async () => {
            delete this.properties[`termlink_${prop.key}`];
            await this._resolveAllSynonyms();
            this.render();
            if (this.context?.propertyPane?.isPropertyPaneOpen()) {
              this.context.propertyPane.refresh();
            }
          }
        })
      );

      if (resolvedNote) {
        dynamicPreFilterFields.push(
          PropertyPaneLabel(`resolved_${prop.key}`, {
            text: resolvedNote
          })
        );
      }

      dynamicPreFilterFields.push(PropertyPaneHorizontalRule());
    });

    return {
      pages: [
        {
          header: {
            description: 'Configure greeting text, search bar, filter chips, background colour, and user profile pre-filtering with term synonyms.'
          },
          groups: [
            {
              groupName: 'Greeting Configuration',
              groupFields: greetingFields
            },
            {
              groupName: 'Search Bar & Display Options',
              groupFields: [
                PropertyPaneTextField('searchPromptText', {
                  label: 'Search bar prompt text',
                  value: props.searchPromptText || 'Search all dashboard cards, tags, metrics...'
                }),
                PropertyPaneToggle('showAppliedFilters', {
                  label: 'Show active filter tags beneath search bar',
                  checked: props.showAppliedFilters !== false
                }),
                PropertyPaneToggle('showAutocomplete', {
                  label: 'Enable autocomplete suggestions',
                  checked: props.showAutocomplete !== false
                }),
                PropertyPaneDropdown('searchAlignment', {
                  label: 'Alignment',
                  options: [
                    { key: 'left', text: 'Left aligned' },
                    { key: 'center', text: 'Centered' },
                    { key: 'right', text: 'Right aligned' }
                  ],
                  selectedKey: props.searchAlignment || 'left'
                })
              ]
            },
            {
              groupName: 'Filter Bar Background Colour',
              groupFields: [
                createColorPickerPropertyField(
                  'backgroundColorField',
                  props.backgroundColor,
                  (color?: string) => {
                    this.properties.backgroundColor = color || 'transparent';
                    this.render();
                  },
                  'Default (transparent - show section background)',
                  'transparent'
                )
              ]
            },
            {
              groupName: 'User Profile Dynamic Pre-Filtering',
              groupFields: dynamicPreFilterFields
            }
          ]
        }
      ]
    };
  }
}


