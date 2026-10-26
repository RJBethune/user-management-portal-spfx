import * as React from 'react';
import { SpfxUiHost } from '@spfx-kit/ui-profile';

interface IPlacement { node: HTMLElement; marker: Comment; }
interface IExpandedState {
  frame: HTMLDivElement;
  placements: IPlacement[];
  background: Map<Element, string | undefined>;
  overflow: { node: HTMLElement; value: string; priority: string }[];
  observer: MutationObserver;
  scroll: { x: number; y: number };
}

/** Move the owned Kit roots together; React, drafts, theme tokens and portals retain their identity. */
export function useExpandedWorkspace(host: SpfxUiHost, className: string): {
  expanded: boolean; native: boolean; pending: boolean; error: string;
  toggle: () => void; enterBrowser: () => void; clearError: () => void;
} {
  const [expanded, setExpanded] = React.useState(false);
  const [native, setNative] = React.useState(false);
  const [pending, setPending] = React.useState(false);
  const [error, setError] = React.useState('');
  const state = React.useRef<IExpandedState>();
  const changing = React.useRef(false);
  const mounted = React.useRef(true);
  const enteredNative = React.useRef(false);
  const document = host.targetDocument;

  const restore = (): void => {
    const current = state.current;
    if (!current) return;
    current.observer.disconnect();
    current.background.forEach((inert, element) => inert === undefined ? element.removeAttribute('inert') : element.setAttribute('inert', inert));
    current.overflow.forEach(({ node, value, priority }) => value ? node.style.setProperty('overflow', value, priority) : node.style.removeProperty('overflow'));
    current.placements.forEach(({ node, marker }) => { marker.parentNode?.insertBefore(node, marker); marker.remove(); });
    current.frame.remove();
    document.defaultView?.scrollTo(current.scroll.x, current.scroll.y);
    state.current = undefined;
    enteredNative.current = false;
    if (mounted.current) {
      setExpanded(false); setNative(false);
      host.appRoot.querySelector<HTMLElement>('[data-workspace-toggle]')?.focus({ preventScroll: true });
    }
  };
  const exit = (): void => {
    if (changing.current) return;
    if (state.current && document.fullscreenElement === state.current.frame) {
      changing.current = true; setPending(true);
      document.exitFullscreen().then(restore, () => {
        if (mounted.current) setError('Use your browser’s full-screen control to exit, then try again.');
      }).then(() => { changing.current = false; if (mounted.current) setPending(false); }).catch(() => undefined);
    } else restore();
  };
  const expand = (): void => {
    if (state.current) return;
    if (document.querySelector('[data-account-expanded]')) { setError('Restore the other expanded Account Management workspace first.'); return; }
    const focused = document.activeElement as HTMLElement | null;
    const scroll = { x: document.defaultView?.scrollX || 0, y: document.defaultView?.scrollY || 0 };
    const frame = document.createElement('div');
    frame.className = className;
    frame.dataset.accountExpanded = host.instanceId;
    frame.setAttribute('role', 'region'); frame.setAttribute('aria-label', 'Expanded Account Management workspace');
    frame.tabIndex = -1;
    const positions = [host.appRoot, ...Array.from(host.appRoot.querySelectorAll('*'))]
      .filter(node => node.scrollTop || node.scrollLeft).map(node => ({ node, top: node.scrollTop, left: node.scrollLeft }));
    const placements = [host.appRoot, host.portalHost].map(node => {
      const marker = document.createComment('Account Management workspace position');
      node.parentNode!.insertBefore(marker, node);
      return { node, marker };
    });
    document.body.appendChild(frame);
    placements.forEach(({ node }) => frame.appendChild(node));
    positions.forEach(({ node, top, left }) => { node.scrollTop = top; node.scrollLeft = left; });
    const background = new Map<Element, string | undefined>();
    const isolate = (): void => {
      Array.from(document.body.children).forEach(child => {
        if (child !== frame && !background.has(child)) { background.set(child, child.getAttribute('inert') ?? undefined); child.setAttribute('inert', ''); }
      });
    };
    isolate();
    const observer = new MutationObserver(isolate);
    observer.observe(document.body, { childList: true });
    const overflow = [document.documentElement, document.body].map(node => ({ node, value: node.style.getPropertyValue('overflow'), priority: node.style.getPropertyPriority('overflow') }));
    overflow.forEach(({ node }) => node.style.setProperty('overflow', 'hidden'));
    state.current = { frame, placements, background, overflow, observer, scroll };
    setExpanded(true); setError('');
    (focused && frame.contains(focused) ? focused : frame).focus({ preventScroll: true });
  };
  React.useLayoutEffect(() => {
    mounted.current = true;
    const keydown = (event: KeyboardEvent): void => {
      const frame = state.current?.frame;
      if (event.key !== 'Escape' || event.defaultPrevented || !frame || !frame.contains(event.target as Node)) return;
      // Menus and confirmations consume Escape first. A second Escape restores the workspace.
      if (frame.querySelector('[role="menu"], [role="dialog"], [role="alertdialog"], [role="listbox"]')) return;
      event.preventDefault(); exit();
    };
    const focus = (event: FocusEvent): void => {
      const frame = state.current?.frame;
      if (frame && !frame.contains(event.target as Node)) (host.appRoot.querySelector<HTMLElement>('[data-workspace-toggle]') || frame).focus({ preventScroll: true });
    };
    const changed = (): void => {
      if (state.current && document.fullscreenElement === state.current.frame) { enteredNative.current = true; setNative(true); }
      else if (enteredNative.current) restore();
    };
    document.addEventListener('keydown', keydown);
    document.addEventListener('focusin', focus);
    document.addEventListener('fullscreenchange', changed);
    return () => {
      mounted.current = false;
      document.removeEventListener('keydown', keydown); document.removeEventListener('focusin', focus); document.removeEventListener('fullscreenchange', changed);
      if (state.current && document.fullscreenElement === state.current.frame) document.exitFullscreen().catch(() => undefined);
      restore();
    };
  }, [host]);
  const enterBrowser = (): void => {
    if (changing.current) return;
    expand();
    const frame = state.current?.frame;
    if (!frame) return;
    const unavailable = 'Browser full screen is unavailable here. The workspace still fills the page.';
    if (!document.fullscreenEnabled || !frame.requestFullscreen || document.fullscreenElement) { setError(unavailable); return; }
    changing.current = true; setPending(true);
    frame.requestFullscreen().catch(() => { if (mounted.current) setError(unavailable); })
      .then(() => { changing.current = false; if (mounted.current) setPending(false); }).catch(() => undefined);
  };
  return { expanded, native, pending, error, toggle: () => { if (!changing.current) { if (state.current) exit(); else expand(); } }, enterBrowser, clearError: () => setError('') };
}
