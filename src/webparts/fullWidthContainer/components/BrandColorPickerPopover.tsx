/**
 * @file BrandColorPickerPopover.tsx
 * @description Visual popup color picker for selecting Turner & Townsend brand colors and custom hex codes.
 * Adheres to Microsoft Fluent UI 2 design standards with interactive swatches, grouping, and active status rings.
 */

import * as React from 'react';
import {
  Button,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Input,
  Caption1,
  makeStyles,
  shorthands,
  tokens,
  Divider,
  FluentProvider,
  webLightTheme
} from '@fluentui/react-components';
import {
  ChevronDownRegular,
  DismissRegular
} from '@fluentui/react-icons';
import { SITE_THEME_COLORS } from './FloatingTextToolbar';

const useStyles = makeStyles({
  triggerBtn: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    width: '100%',
    padding: '6px 10px',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    cursor: 'pointer',
    outlineStyle: 'none',
    boxSizing: 'border-box',
    transitionProperty: 'border-color, background-color, box-shadow',
    transitionDuration: '120ms',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    },
    ':focus': {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      boxShadow: tokens.shadow4
    }
  },
  triggerLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    overflow: 'hidden'
  },
  triggerPreviewSwatch: {
    width: '20px',
    height: '20px',
    ...shorthands.borderRadius('4px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    flexShrink: 0
  },
  triggerText: {
    fontSize: '0.85rem',
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground1,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  popoverSurface: {
    padding: '10px',
    minWidth: '250px',
    maxWidth: '280px',
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    boxShadow: tokens.shadow28,
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: '#FFFFFF !important' as any,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1)
  },
  headerRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingBottom: '4px'
  },
  sectionTitle: {
    fontSize: '0.72rem',
    fontWeight: 600,
    color: tokens.colorNeutralForeground3,
    padding: '4px 2px 2px 2px',
    textTransform: 'uppercase',
    letterSpacing: '0.5px'
  },
  colorGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(6, 1fr)',
    gap: '6px',
    marginBottom: '4px'
  },
  colorSwatch: {
    width: '24px',
    height: '24px',
    ...shorthands.borderRadius('4px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    cursor: 'pointer',
    boxSizing: 'border-box',
    transitionProperty: 'transform, border-color, box-shadow',
    transitionDuration: '100ms',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    ':hover': {
      transform: 'scale(1.15)',
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      boxShadow: tokens.shadow4,
      zIndex: 10
    }
  },
  selectedSwatch: {
    outline: `2px solid ${tokens.colorBrandStroke1}`,
    outlineOffset: '1px'
  },
  defaultOptionRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: '6px 8px',
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    cursor: 'pointer',
    marginBottom: '4px',
    transitionProperty: 'background-color, border-color, box-shadow',
    transitionDuration: '100ms',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground2Hover,
      ...shorthands.borderColor(tokens.colorBrandStroke1)
    }
  },
  defaultOptionRowSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    ...shorthands.borderColor(tokens.colorBrandStroke1),
    boxShadow: `0 0 0 1px ${tokens.colorBrandStroke1}`
  },
  defaultSwatchPreview: {
    width: '20px',
    height: '20px',
    ...shorthands.borderRadius('4px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    flexShrink: 0,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },
  customInputRow: {
    display: 'flex',
    gap: '6px',
    alignItems: 'center',
    marginTop: '6px'
  }
});

export interface IBrandColorPickerPopoverProps {
  selectedColor?: string;
  onChange: (color?: string) => void;
  defaultLabel?: string;
  defaultColorHex?: string;
  disabled?: boolean;
}

/**
 * Visual Popover component for Turner & Townsend brand and custom background color selection.
 */
export const BrandColorPickerPopover: React.FC<IBrandColorPickerPopoverProps> = ({
  selectedColor,
  onChange,
  defaultLabel = 'Default (inherit background)',
  defaultColorHex = 'transparent',
  disabled = false
}) => {
  const styles = useStyles();
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const [customHex, setCustomHex] = React.useState<string>(selectedColor || '');

  React.useEffect(() => {
    setCustomHex(selectedColor || '');
  }, [selectedColor]);

  // Find color friendly name if from palette
  const selectedColorName = React.useMemo(() => {
    if (!selectedColor) return defaultLabel;
    const match = SITE_THEME_COLORS.find(
      (c) => c.hex.toLowerCase() === selectedColor.toLowerCase()
    );
    return match ? `${match.name} (${match.hex})` : selectedColor;
  }, [selectedColor, defaultLabel]);

  const handleSelect = (hex?: string): void => {
    onChange(hex);
    setIsOpen(false);
  };

  const effectiveColor = selectedColor || defaultColorHex;
  const isEffectiveTransparent = !effectiveColor || effectiveColor === 'transparent';
  const isDefaultSelected = !selectedColor;

  return (
    <Popover
      open={isOpen}
      onOpenChange={(e, data) => setIsOpen(data.open)}
      positioning="below-start"
    >
      <PopoverTrigger disableButtonEnhancement>
        <button
          type="button"
          className={styles.triggerBtn}
          disabled={disabled}
          title={selectedColorName}
          onClick={() => setIsOpen(!isOpen)}
        >
          <div className={styles.triggerLeft}>
            <div
              className={styles.triggerPreviewSwatch}
              style={{
                backgroundColor: !isEffectiveTransparent ? effectiveColor : 'transparent',
                backgroundImage: !isEffectiveTransparent
                  ? 'none'
                  : 'linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0), linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0)',
                backgroundSize: '8px 8px',
                backgroundPosition: '0 0, 4px 4px'
              }}
            />
            <span className={styles.triggerText}>{selectedColorName}</span>
          </div>
          <ChevronDownRegular style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
        </button>
      </PopoverTrigger>

      <PopoverSurface className={styles.popoverSurface}>
        <FluentProvider theme={webLightTheme}>
          <div className={styles.headerRow}>
          <Caption1 style={{ fontWeight: 600, color: tokens.colorNeutralForeground1 }}>
            Brand Colour Palette
          </Caption1>
          {selectedColor && (
            <Button
              size="small"
              appearance="subtle"
              icon={<DismissRegular />}
              onClick={() => handleSelect(undefined)}
              title="Reset to default background"
              style={{ fontSize: '0.72rem', height: '22px', padding: '0 4px' }}
            >
              Reset
            </Button>
          )}
        </div>

        <Divider style={{ margin: '2px 0 6px 0' }} />

        {/* Dedicated Contextual Default Option Swatch */}
        <div className={styles.sectionTitle}>Default Option</div>
        <div
          className={`${styles.defaultOptionRow} ${isDefaultSelected ? styles.defaultOptionRowSelected : ''}`}
          onClick={() => handleSelect(undefined)}
          title={`Select ${defaultLabel} (${defaultColorHex})`}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' }}>
            <div
              className={styles.defaultSwatchPreview}
              style={{
                backgroundColor: defaultColorHex !== 'transparent' ? defaultColorHex : 'transparent',
                backgroundImage: defaultColorHex === 'transparent'
                  ? 'linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0), linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0)'
                  : undefined,
                backgroundSize: '6px 6px',
                backgroundPosition: '0 0, 3px 3px',
                borderColor: (defaultColorHex === '#FFFFFF' || defaultColorHex === 'transparent') ? '#d1d1d1' : undefined
              }}
            />
            <span style={{ fontSize: '0.8rem', fontWeight: isDefaultSelected ? 600 : 400, color: tokens.colorNeutralForeground1, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
              {defaultLabel}
            </span>
          </div>
          {isDefaultSelected && (
            <Caption1 style={{ fontSize: '0.72rem', fontWeight: 600, color: tokens.colorBrandForeground1 }}>
              ✓ Active
            </Caption1>
          )}
        </div>

        <Divider style={{ margin: '4px 0 6px 0' }} />

        {/* TT Primary */}
        <div className={styles.sectionTitle}>TT Primary (Blue, Cyan, Grey)</div>
        <div className={styles.colorGrid}>
          {SITE_THEME_COLORS.filter((c) => c.group === 'primary').map((c) => {
            const isSelected = selectedColor?.toLowerCase() === c.hex.toLowerCase();
            return (
              <div
                key={c.name}
                className={`${styles.colorSwatch} ${isSelected ? styles.selectedSwatch : ''}`}
                style={{ backgroundColor: c.hex }}
                title={`${c.name} (${c.hex})`}
                onClick={() => handleSelect(c.hex)}
              />
            );
          })}
        </div>

        {/* TT Secondary */}
        <div className={styles.sectionTitle}>TT Secondary (Green, Orange)</div>
        <div className={styles.colorGrid}>
          {SITE_THEME_COLORS.filter((c) => c.group === 'secondary').map((c) => {
            const isSelected = selectedColor?.toLowerCase() === c.hex.toLowerCase();
            return (
              <div
                key={c.name}
                className={`${styles.colorSwatch} ${isSelected ? styles.selectedSwatch : ''}`}
                style={{ backgroundColor: c.hex }}
                title={`${c.name} (${c.hex})`}
                onClick={() => handleSelect(c.hex)}
              />
            );
          })}
        </div>

        {/* Backgrounds & Neutrals */}
        <div className={styles.sectionTitle}>Backgrounds &amp; Neutrals</div>
        <div className={styles.colorGrid}>
          {SITE_THEME_COLORS.filter((c) => c.group === 'backgrounds').map((c) => {
            const isSelected = selectedColor?.toLowerCase() === c.hex.toLowerCase();
            const isTransparent = c.hex === 'transparent';
            return (
              <div
                key={c.name}
                className={`${styles.colorSwatch} ${isSelected ? styles.selectedSwatch : ''}`}
                style={{
                  backgroundColor: c.hex,
                  backgroundImage: isTransparent
                    ? 'linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0), linear-gradient(45deg, #e0e0e0 25%, transparent 25%, transparent 75%, #e0e0e0 75%, #e0e0e0)'
                    : undefined,
                  backgroundSize: isTransparent ? '6px 6px' : undefined,
                  backgroundPosition: isTransparent ? '0 0, 3px 3px' : undefined,
                  borderColor: (c.hex === '#FFFFFF' || isTransparent) ? '#d1d1d1' : undefined
                }}
                title={`${c.name} (${c.hex})`}
                onClick={() => handleSelect(c.hex)}
              />
            );
          })}
        </div>

        {/* Brand Tints */}
        <div className={styles.sectionTitle}>Brand Tints (80%, 60%, 40%, 20%)</div>
        <div className={styles.colorGrid}>
          {SITE_THEME_COLORS.filter((c) => c.group === 'tints').map((c) => {
            const isSelected = selectedColor?.toLowerCase() === c.hex.toLowerCase();
            return (
              <div
                key={c.name}
                className={`${styles.colorSwatch} ${isSelected ? styles.selectedSwatch : ''}`}
                style={{ backgroundColor: c.hex }}
                title={`${c.name} (${c.hex})`}
                onClick={() => handleSelect(c.hex)}
              />
            );
          })}
        </div>

        <Divider style={{ margin: '6px 0 4px 0' }} />

        {/* Custom Hex Input */}
        <div className={styles.customInputRow}>
          <Input
            size="small"
            placeholder="#HEX (e.g. #F2EEE7)"
            value={customHex}
            onChange={(e, data) => setCustomHex(data.value)}
            style={{ flex: 1 }}
          />
          <Button
            size="small"
            appearance="primary"
            onClick={() => handleSelect(customHex.trim() ? customHex.trim() : undefined)}
          >
            Apply
          </Button>
        </div>
        </FluentProvider>
      </PopoverSurface>
    </Popover>
  );
};
