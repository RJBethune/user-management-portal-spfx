import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SpfxUiHost } from '@spfx-kit/ui-profile';
import { applyCompiledThemeArtifact, applySharePointSvelteLayer, parseCompiledThemeArtifact, resolveSharePointSiteTheme, sharePointSvelteLayer } from '@spfx-kit/ui-theme/runtime';
import snapshot from './kit.compiled.json';

const artifact = parseCompiledThemeArtifact(snapshot);
export type ThemeChoice = 'sharepoint' | 'light' | 'dark';
export const initialTokens = artifact.compiled.light.tokens;

export function applyAppTheme(host: SpfxUiHost, choice: ThemeChoice, site?: IReadonlyTheme): void {
  const layer = choice === 'sharepoint' ? sharePointSvelteLayer(resolveSharePointSiteTheme(site || {})) : undefined;
  const mode: 'light' | 'dark' = layer && layer.mode ? layer.mode : choice === 'dark' ? 'dark' : 'light';
  applyCompiledThemeArtifact(host, artifact, mode, artifact.compatibility);
  if (layer) applySharePointSvelteLayer(host, artifact.compiled[mode].tokens, layer);
}

export function readThemeChoice(key: string): ThemeChoice {
  try {
    const value = localStorage.getItem(key);
    return value === 'light' || value === 'dark' ? value : 'sharepoint';
  } catch { return 'sharepoint'; }
}
