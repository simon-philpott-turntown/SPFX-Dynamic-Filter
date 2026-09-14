/**
 * @file FloatingTextToolbar.tsx
 * @description Floating rich text formatting toolbar modeled after the native SharePoint Text Web Part ribbon.
 * Features Text Style (H1, H2, H3, H4, Body), Font Family selector, comprehensive SharePoint Site Theme Palette,
 * highlight backgrounds, list formatting, and text alignment.
 */

import * as React from 'react';
import {
  Button,
  makeStyles,
  shorthands,
  tokens,
  Divider,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Caption1
} from '@fluentui/react-components';
import {
  TextBoldRegular,
  TextItalicRegular,
  TextUnderlineRegular,
  TextColorRegular,
  HighlightRegular,
  LinkRegular,
  TextClearFormattingRegular,
  TextAlignLeftRegular,
  TextAlignCenterRegular,
  TextAlignRightRegular,
  TextBulletListRegular,
  TextNumberListLtrRegular,
  LineHorizontal1Regular,
  TextStrikethroughRegular,
  CodeRegular,
  TextFontRegular,
  TextHeader1Regular,
  ChevronDownRegular
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  toolbarContainer: {
    display: 'inline-flex',
    alignItems: 'center',
    ...shorthands.padding('4px', '8px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow28,
    ...shorthands.gap('3px'),
    zIndex: 2000,
    marginBottom: tokens.spacingVerticalXS,
    flexWrap: 'nowrap',
    whiteSpace: 'nowrap',
    width: 'max-content',
    minWidth: 'max-content',
    maxWidth: 'none',
    userSelect: 'none'
  },
  dragHandle: {
    cursor: 'grab',
    display: 'flex',
    alignItems: 'center',
    color: tokens.colorNeutralForeground3,
    paddingRight: '6px',
    paddingLeft: '2px',
    transitionProperty: 'color',
    transitionDuration: '100ms',
    ':hover': {
      color: tokens.colorNeutralForeground1
    }
  },
  btnActive: {
    backgroundColor: tokens.colorNeutralBackground1Selected,
    color: tokens.colorBrandForeground1,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorBrandForeground1
    }
  },
  dropdownBtn: {
    height: '28px',
    fontSize: '0.8rem',
    fontWeight: tokens.fontWeightSemibold,
    padding: '0 8px',
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover
    }
  },
  menuSurface: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    padding: '6px',
    minWidth: '160px',
    maxHeight: '260px',
    overflowY: 'auto',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow28
  },
  menuItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '6px 10px',
    borderRadius: tokens.borderRadiusSmall,
    cursor: 'pointer',
    fontSize: '0.85rem',
    transitionProperty: 'background-color, color',
    transitionDuration: '100ms',
    ':hover': {
      backgroundColor: tokens.colorBrandBackground2,
      color: tokens.colorBrandForeground1
    }
  },
  colorPickerSurface: {
    padding: '10px',
    minWidth: '240px',
    maxWidth: '270px',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow28
  },
  highlightPickerSurface: {
    padding: '8px',
    minWidth: '190px',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    boxShadow: tokens.shadow28
  },
  colorGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(6, 1fr)',
    ...shorthands.gap('6px'),
    ...shorthands.padding('6px', '8px')
  },
  colorSectionHeader: {
    fontSize: '0.72rem',
    fontWeight: 600,
    color: tokens.colorNeutralForeground3,
    padding: '4px 8px 2px 8px',
    textTransform: 'uppercase',
    letterSpacing: '0.5px'
  },
  colorSwatch: {
    width: '22px',
    height: '22px',
    ...shorthands.borderRadius('4px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    cursor: 'pointer',
    transitionProperty: 'transform, border-color, box-shadow',
    transitionDuration: '100ms',
    ':hover': {
      transform: 'scale(1.18)',
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      boxShadow: tokens.shadow4
    }
  }
});

export interface IFloatingTextToolbarProps {
  onFormat?: (command: string, value?: string) => void;
  isBold?: boolean;
  isItalic?: boolean;
  isUnderline?: boolean;
  isStrikethrough?: boolean;
}

/**
 * Turner & Townsend Official Brand Color Palette (Version 2.2 — July 2025)
 */
export const SITE_THEME_COLORS = [
  // Primary Brand Colors & 100+ Shades
  { name: 'TT Blue (Main)', hex: '#1E4479', group: 'primary' },
  { name: 'TT Blue (100+ Deep)', hex: '#001436', group: 'primary' },
  { name: 'TT Cyan (Main)', hex: '#0090DC', group: 'primary' },
  { name: 'TT Cyan (100+ Deep)', hex: '#0073A5', group: 'primary' },
  { name: 'TT Grey (Main)', hex: '#505A60', group: 'primary' },
  { name: 'TT Grey (100+ Dark)', hex: '#292929', group: 'primary' },

  // Secondary Brand Colors & 100+ Shades
  { name: 'TT Green (Main)', hex: '#00A000', group: 'secondary' },
  { name: 'TT Green (100+ Deep)', hex: '#007E1E', group: 'secondary' },
  { name: 'TT Orange (Main)', hex: '#D55C17', group: 'secondary' },
  { name: 'TT Orange (100+ Deep)', hex: '#C83700', group: 'secondary' },

  // Backgrounds & Neutrals
  { name: 'Transparent (Show Through)', hex: 'transparent', group: 'backgrounds' },
  { name: 'TT Mushroom (Background)', hex: '#F2EEE7', group: 'backgrounds' },
  { name: 'TT White', hex: '#FFFFFF', group: 'backgrounds' },
  { name: 'Dark Neutral', hex: '#1F1F1F', group: 'backgrounds' },
  { name: 'Mid Neutral', hex: '#616161', group: 'backgrounds' },

  // Primary & Secondary Brand Tints
  { name: 'TT Blue (80%)', hex: '#4B6994', group: 'tints' },
  { name: 'TT Blue (60%)', hex: '#788FAE', group: 'tints' },
  { name: 'TT Blue (40%)', hex: '#A5B4C9', group: 'tints' },
  { name: 'TT Blue (20%)', hex: '#D2DAE4', group: 'tints' },

  { name: 'TT Cyan (80%)', hex: '#33A6E3', group: 'tints' },
  { name: 'TT Cyan (60%)', hex: '#66BCEB', group: 'tints' },
  { name: 'TT Cyan (40%)', hex: '#99D3F1', group: 'tints' },
  { name: 'TT Cyan (20%)', hex: '#CCE9F8', group: 'tints' },

  { name: 'TT Grey (80%)', hex: '#737B80', group: 'tints' },
  { name: 'TT Grey (60%)', hex: '#969CA0', group: 'tints' },
  { name: 'TT Grey (40%)', hex: '#B9BDC0', group: 'tints' },
  { name: 'TT Grey (20%)', hex: '#DCDFE0', group: 'tints' },

  { name: 'TT Green (80%)', hex: '#33B333', group: 'tints' },
  { name: 'TT Green (60%)', hex: '#66C666', group: 'tints' },
  { name: 'TT Green (40%)', hex: '#99D999', group: 'tints' },
  { name: 'TT Green (20%)', hex: '#CCEECC', group: 'tints' },

  { name: 'TT Orange (80%)', hex: '#DD7D45', group: 'tints' },
  { name: 'TT Orange (60%)', hex: '#E69D74', group: 'tints' },
  { name: 'TT Orange (40%)', hex: '#EEBEA2', group: 'tints' },
  { name: 'TT Orange (20%)', hex: '#F7DED1', group: 'tints' }
];

export const HIGHLIGHT_COLORS = [
  { name: 'Bright Yellow', hex: '#fff100' },
  { name: 'Pastel Green', hex: '#92c353' },
  { name: 'Sky Cyan', hex: '#00bcf2' },
  { name: 'Soft Lavender', hex: '#b4a0ff' },
  { name: 'Warm Peach', hex: '#ffb900' },
  { name: 'Light Rose', hex: '#e3008c' }
];

export const FONT_FAMILIES = [
  { name: 'Segoe UI (Standard)', value: 'Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, sans-serif' },
  { name: 'Aptos (Modern 365)', value: 'Aptos, Segoe UI, sans-serif' },
  { name: 'Arial', value: 'Arial, Helvetica, sans-serif' },
  { name: 'Calibri', value: 'Calibri, Candara, Segoe, sans-serif' },
  { name: 'Georgia', value: 'Georgia, Cambria, serif' },
  { name: 'Tahoma', value: 'Tahoma, Verdana, Segoe, sans-serif' },
  { name: 'Trebuchet MS', value: 'Trebuchet MS, Lucida Sans Unicode, sans-serif' },
  { name: 'Times New Roman', value: 'Times New Roman, Times, serif' },
  { name: 'Courier New (Code)', value: 'Courier New, Courier, monospace' }
];

export const TEXT_STYLES = [
  { label: 'Header 1', tag: '<h1>', style: { fontSize: '1.75rem', fontWeight: 700 } },
  { label: 'Header 2', tag: '<h2>', style: { fontSize: '1.35rem', fontWeight: 600 } },
  { label: 'Header 3', tag: '<h3>', style: { fontSize: '1.15rem', fontWeight: 600 } },
  { label: 'Header 4', tag: '<h4>', style: { fontSize: '1rem', fontWeight: 600 } },
  { label: 'Body text', tag: '<p>', style: { fontSize: '0.9rem', fontWeight: 400 } },
  { label: 'Quote', tag: '<blockquote>', style: { fontStyle: 'italic', borderLeft: '3px solid #0078d4', paddingLeft: '8px' } },
  { label: 'Code block', tag: '<pre>', style: { fontFamily: 'monospace', fontSize: '0.85rem' } }
];

export const FloatingTextToolbar: React.FC<IFloatingTextToolbarProps> = ({
  onFormat,
  isBold = false,
  isItalic = false,
  isUnderline = false,
  isStrikethrough = false
}) => {
  const styles = useStyles();
  const [selectedStyleLabel, setSelectedStyleLabel] = React.useState<string>('Style');
  const [selectedFontLabel, setSelectedFontLabel] = React.useState<string>('Font');

  const handleCommand = (cmd: string, val?: string): void => {
    if (onFormat) {
      onFormat(cmd, val);
    } else {
      document.execCommand(cmd, false, val);
    }
  };

  /**
   * Prevents mouse down default action so text selection inside contentEditable is NOT blurred or cleared.
   */
  const handleToolbarMouseDown = (e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
  };

  return (
    <div
      className={`${styles.toolbarContainer} floating-toolbar-container fui-floating-toolbar`}
      onMouseDown={handleToolbarMouseDown}
      onClick={(e) => e.stopPropagation()}
    >
      {/* Drag handle */}
      <div className={styles.dragHandle} title="Formatting Ribbon">
        <LineHorizontal1Regular fontSize={14} />
      </div>

      {/* Text Style / Type Dropdown (H1, H2, H3, H4, Body) */}
      <Popover positioning="below" inline>
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            className={styles.dropdownBtn}
            icon={<TextHeader1Regular />}
            title="Text Style (Heading 1, 2, 3, Body)"
            onMouseDown={handleToolbarMouseDown}
          >
            {selectedStyleLabel}
            <ChevronDownRegular fontSize={10} />
          </Button>
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} className={`${styles.menuSurface} floating-toolbar-container fui-floating-toolbar`}>
          <Caption1 style={{ fontWeight: 600, padding: '2px 8px', color: tokens.colorNeutralForeground3 }}>
            Text Styles
          </Caption1>
          {TEXT_STYLES.map((st) => (
            <div
              key={st.label}
              className={styles.menuItem}
              onMouseDown={handleToolbarMouseDown}
              onClick={() => {
                handleCommand('formatBlock', st.tag);
                setSelectedStyleLabel(st.label);
              }}
            >
              <span style={st.style}>{st.label}</span>
            </div>
          ))}
        </PopoverSurface>
      </Popover>

      {/* Font Family Selector Dropdown */}
      <Popover positioning="below" inline>
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            className={styles.dropdownBtn}
            icon={<TextFontRegular />}
            title="Font Family"
            onMouseDown={handleToolbarMouseDown}
          >
            {selectedFontLabel}
            <ChevronDownRegular fontSize={10} />
          </Button>
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} className={`${styles.menuSurface} floating-toolbar-container fui-floating-toolbar`}>
          <Caption1 style={{ fontWeight: 600, padding: '2px 8px', color: tokens.colorNeutralForeground3 }}>
            Font Families
          </Caption1>
          {FONT_FAMILIES.map((font) => (
            <div
              key={font.name}
              className={styles.menuItem}
              onMouseDown={handleToolbarMouseDown}
              onClick={() => {
                handleCommand('fontName', font.value);
                setSelectedFontLabel(font.name.split(' ')[0]);
              }}
            >
              <span style={{ fontFamily: font.value }}>{font.name}</span>
            </div>
          ))}
        </PopoverSurface>
      </Popover>

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

      {/* Bold, Italic, Underline, Strikethrough */}
      <Button
        size="small"
        appearance="subtle"
        className={isBold ? styles.btnActive : ''}
        icon={<TextBoldRegular />}
        title="Bold (Ctrl+B)"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('bold')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isItalic ? styles.btnActive : ''}
        icon={<TextItalicRegular />}
        title="Italic (Ctrl+I)"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('italic')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isUnderline ? styles.btnActive : ''}
        icon={<TextUnderlineRegular />}
        title="Underline (Ctrl+U)"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('underline')}
      />
      <Button
        size="small"
        appearance="subtle"
        className={isStrikethrough ? styles.btnActive : ''}
        icon={<TextStrikethroughRegular />}
        title="Strikethrough"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('strikeThrough')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<CodeRegular />}
        title="Inline code snippet"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('formatBlock', '<pre>')}
      />

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

      {/* Full SharePoint Theme Color Picker Popover */}
      <Popover positioning="below" inline>
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            icon={<TextColorRegular />}
            title="SharePoint Site Theme Colour Palette"
            onMouseDown={handleToolbarMouseDown}
          />
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} className={`${styles.colorPickerSurface} floating-toolbar-container fui-floating-toolbar`}>
          {/* Default Text Color Option */}
          <div className={styles.colorSectionHeader}>Default Option</div>
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              padding: '4px 8px',
              backgroundColor: tokens.colorNeutralBackground2,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: '4px',
              cursor: 'pointer',
              marginBottom: '6px',
              margin: '0 8px 6px 8px'
            }}
            onMouseDown={handleToolbarMouseDown}
            onClick={() => handleCommand('foreColor', '#242424')}
            title="Reset to default text color (inherit / dark neutral #242424)"
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <div
                style={{
                  width: '18px',
                  height: '18px',
                  borderRadius: '3px',
                  border: `1px solid ${tokens.colorNeutralStroke1}`,
                  backgroundColor: '#242424'
                }}
              />
              <span style={{ fontSize: '0.78rem', color: tokens.colorNeutralForeground1, fontWeight: 500 }}>
                Default (Inherit / #242424)
              </span>
            </div>
          </div>

          <Divider style={{ margin: '2px 8px 6px 8px' }} />

          <div className={styles.colorSectionHeader}>TT Primary (Blue, Cyan, Grey)</div>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.filter((c) => c.group === 'primary').map((c) => (
              <div
                key={c.name}
                className={styles.colorSwatch}
                style={{ backgroundColor: c.hex }}
                title={`${c.name} (${c.hex})`}
                onMouseDown={handleToolbarMouseDown}
                onClick={() => handleCommand('foreColor', c.hex)}
              />
            ))}
          </div>

          <div className={styles.colorSectionHeader}>TT Secondary (Green, Orange)</div>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.filter((c) => c.group === 'secondary').map((c) => (
              <div
                key={c.name}
                className={styles.colorSwatch}
                style={{ backgroundColor: c.hex }}
                title={`${c.name} (${c.hex})`}
                onMouseDown={handleToolbarMouseDown}
                onClick={() => handleCommand('foreColor', c.hex)}
              />
            ))}
          </div>

          <div className={styles.colorSectionHeader}>Backgrounds & Neutrals</div>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.filter((c) => c.group === 'backgrounds').map((c) => (
              <div
                key={c.name}
                className={styles.colorSwatch}
                style={{ backgroundColor: c.hex, border: c.hex === '#FFFFFF' ? '1px solid #d1d1d1' : undefined }}
                title={`${c.name} (${c.hex})`}
                onMouseDown={handleToolbarMouseDown}
                onClick={() => handleCommand('foreColor', c.hex)}
              />
            ))}
          </div>

          <div className={styles.colorSectionHeader}>Brand Tints (80%, 60%, 40%, 20%)</div>
          <div className={styles.colorGrid}>
            {SITE_THEME_COLORS.filter((c) => c.group === 'tints').map((c) => (
              <div
                key={c.name}
                className={styles.colorSwatch}
                style={{ backgroundColor: c.hex }}
                title={`${c.name} (${c.hex})`}
                onMouseDown={handleToolbarMouseDown}
                onClick={() => handleCommand('foreColor', c.hex)}
              />
            ))}
          </div>
        </PopoverSurface>
      </Popover>

      {/* Highlight Background Color Picker Popover */}
      <Popover positioning="below" inline>
        <PopoverTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            icon={<HighlightRegular />}
            title="Highlight Background Colour"
            onMouseDown={handleToolbarMouseDown}
          />
        </PopoverTrigger>
        <PopoverSurface onMouseDown={handleToolbarMouseDown} className={`${styles.highlightPickerSurface} floating-toolbar-container fui-floating-toolbar`}>
          {/* Default Highlight Option */}
          <div className={styles.colorSectionHeader}>Default Option</div>
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              padding: '4px 8px',
              backgroundColor: tokens.colorNeutralBackground2,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: '4px',
              cursor: 'pointer',
              marginBottom: '6px',
              margin: '0 8px 6px 8px'
            }}
            onMouseDown={handleToolbarMouseDown}
            onClick={() => handleCommand('hiliteColor', 'transparent')}
            title="No highlight background (transparent)"
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <div
                style={{
                  width: '18px',
                  height: '18px',
                  borderRadius: '3px',
                  border: `1px solid ${tokens.colorNeutralStroke1}`,
                  backgroundColor: 'transparent',
                  backgroundImage: 'linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0), linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0)',
                  backgroundSize: '6px 6px',
                  backgroundPosition: '0 0, 3px 3px'
                }}
              />
              <span style={{ fontSize: '0.78rem', color: tokens.colorNeutralForeground1, fontWeight: 500 }}>
                None (Transparent)
              </span>
            </div>
          </div>

          <Divider style={{ margin: '2px 8px 6px 8px' }} />

          <div className={styles.colorSectionHeader}>Highlight Backgrounds</div>
          <div className={styles.colorGrid}>
            {HIGHLIGHT_COLORS.map((c) => (
              <div
                key={c.name}
                className={styles.colorSwatch}
                style={{ backgroundColor: c.hex }}
                title={c.name}
                onMouseDown={handleToolbarMouseDown}
                onClick={() => handleCommand('hiliteColor', c.hex)}
              />
            ))}
          </div>
        </PopoverSurface>
      </Popover>

      {/* Insert Link */}
      <Button
        size="small"
        appearance="subtle"
        icon={<LinkRegular />}
        title="Insert Hyperlink"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => {
          const url = window.prompt('Enter destination URL:');
          if (url) handleCommand('createLink', url);
        }}
      />

      {/* Clear Formatting */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextClearFormattingRegular />}
        title="Clear Formatting"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('removeFormat')}
      />

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

      {/* Alignment */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignLeftRegular />}
        title="Align Left"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('justifyLeft')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignCenterRegular />}
        title="Align Centre"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('justifyCenter')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextAlignRightRegular />}
        title="Align Right"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('justifyRight')}
      />

      <Divider vertical style={{ height: '18px', margin: '0 2px' }} />

      {/* Lists */}
      <Button
        size="small"
        appearance="subtle"
        icon={<TextBulletListRegular />}
        title="Bullet List"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('insertUnorderedList')}
      />
      <Button
        size="small"
        appearance="subtle"
        icon={<TextNumberListLtrRegular />}
        title="Numbered List"
        onMouseDown={handleToolbarMouseDown}
        onClick={() => handleCommand('insertOrderedList')}
      />
    </div>
  );
};
