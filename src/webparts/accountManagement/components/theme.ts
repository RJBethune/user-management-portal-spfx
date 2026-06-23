import { Theme, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

/**
 * Maps the SharePoint site theme onto Fluent UI v9 brand tokens so the web part
 * follows org branding and section backgrounds (including inverted/dark sections).
 * Pass the section theme (IReadonlyTheme) when available; with no argument it
 * returns the default light theme. (Pattern shared with task-tracker-spfx.)
 */
export const buildFluentTheme = (sp?: IReadonlyTheme): Theme => {
  const base: Theme = sp && sp.isInverted ? webDarkTheme : webLightTheme;
  const p = sp ? sp.palette : undefined;
  if (!p || !p.themePrimary) { return base; }

  const primary: string = p.themePrimary;
  const darkAlt: string = p.themeDarkAlt || primary;
  const dark: string = p.themeDark || primary;
  const light: string = p.themeLight || primary;

  return {
    ...base,
    colorBrandBackground: primary,
    colorBrandBackgroundHover: darkAlt,
    colorBrandBackgroundPressed: dark,
    colorBrandBackgroundSelected: dark,
    colorBrandForeground1: primary,
    colorBrandForeground2: darkAlt,
    colorBrandForegroundLink: primary,
    colorBrandForegroundLinkHover: darkAlt,
    colorBrandForegroundLinkPressed: dark,
    colorBrandStroke1: primary,
    colorBrandStroke2: light,
    colorCompoundBrandBackground: primary,
    colorCompoundBrandBackgroundHover: darkAlt,
    colorCompoundBrandBackgroundPressed: dark,
    colorCompoundBrandStroke: primary,
    colorCompoundBrandStrokeHover: darkAlt,
    colorCompoundBrandForeground1: primary,
    colorNeutralForegroundOnBrand: p.white || '#ffffff'
  };
};
