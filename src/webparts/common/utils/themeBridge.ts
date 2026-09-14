/**
 * @file themeBridge.ts
 * @description Bridges SharePoint SPFx IReadonlyTheme / semanticColors to Fluent UI 2 (v9) Theme tokens.
 * Adheres strictly to the Fluent 2 Design System (https://fluent2.microsoft.design/).
 * Placed in common/utils so it can be imported by any web part bundle without cross-bundle dependencies.
 */

import {
  Theme,
  webLightTheme,
  webDarkTheme
} from '@fluentui/react-components';

export interface ISharePointThemeColors {
  themePrimary?: string;
  themeLighterAlt?: string;
  themeLighter?: string;
  themeLight?: string;
  themeTertiary?: string;
  themeSecondary?: string;
  themeDarkAlt?: string;
  themeDark?: string;
  themeDarker?: string;
  neutralLighterAlt?: string;
  neutralLighter?: string;
  neutralLight?: string;
  neutralQuaternaryAlt?: string;
  neutralQuaternary?: string;
  neutralTertiaryAlt?: string;
  neutralTertiary?: string;
  neutralSecondary?: string;
  neutralPrimaryAlt?: string;
  neutralPrimary?: string;
  neutralDark?: string;
  black?: string;
  white?: string;
  bodyBackground?: string;
  bodyText?: string;
}

/**
 * Creates a Fluent 2 Theme inheriting the active SharePoint Online site palette.
 * @param spTheme - The IReadonlyTheme or palette colors provided by SPFx
 * @param isDark - Whether the current site theme is dark/inverted
 * @returns Fluent 2 Theme object for <FluentProvider>
 */
export function getFluent2Theme(spTheme?: any, isDark?: boolean): Theme {
  const baseTheme = isDark ? webDarkTheme : webLightTheme;

  if (!spTheme) {
    return baseTheme;
  }

  const palette: ISharePointThemeColors = spTheme.palette || spTheme;
  const semanticColors = spTheme.semanticColors || {};

  const brandPrimary = palette.themePrimary || semanticColors.primaryButtonBackground;

  if (!brandPrimary) {
    return baseTheme;
  }

  return {
    ...baseTheme,
    colorBrandForeground1: brandPrimary,
    colorBrandForeground2: palette.themeDarkAlt || brandPrimary,
    colorBrandBackground: brandPrimary,
    colorBrandBackgroundHover: palette.themeDarkAlt || brandPrimary,
    colorBrandBackgroundPressed: palette.themeDarker || brandPrimary,
    colorBrandStroke1: brandPrimary,
    colorBrandStroke2: palette.themeLight || brandPrimary
  };
}
