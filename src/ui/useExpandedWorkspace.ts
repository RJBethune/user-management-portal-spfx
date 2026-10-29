import * as React from 'react';
import { SpfxUiHost } from '@spfx-kit/ui-profile';
import { elevateWorkspace, WorkspaceLayer } from '@spfx-kit/ui-patterns/layout-stability';
import { animateValue, spring } from 'motion-dom';

type Bounds = { left: number; top: number; width: number; height: number };
interface IPlacement { node: HTMLElement; marker: Comment; }
interface IExpandedState {
  frame: HTMLDivElement;
  placeholder: HTMLDivElement;
  placements: IPlacement[];
  layer: WorkspaceLayer;
  background: Map<Element, string | undefined>;
  overflow: { node: HTMLElement; value: string; priority: string }[];
  observer: MutationObserver;
  scroll: { x: number; y: number };
}
const autoExpandedDocuments = new WeakSet<Document>();

/** Same mounted Kit roots, Better-family spring and Kit top layer; drafts never remount. */
export function useExpandedWorkspace(host: SpfxUiHost, className: string): {
  expanded: boolean; native: boolean; pending: boolean; error: string;
  toggle: () => void; enterBrowser: () => void; clearError: () => void;
} {
  const [expanded, setExpanded] = React.useState(false);
  const [native, setNative] = React.useState(false);
  const [pending, setPending] = React.useState(false);
  const [error, setError] = React.useState('');
  const controls = React.useRef<{ toggle: () => void; enterBrowser: () => void }>();

  React.useLayoutEffect(() => {
    const doc = host.targetDocument;
    const win = doc.defaultView!;
    const motion = win.matchMedia('(prefers-reduced-motion: reduce)');
    let state: IExpandedState | undefined;
    let animation: ReturnType<typeof animateValue> | undefined;
    let nextFrame = 0;
    let targetExpanded = false;
    let changingNative = false;
    let enteredNative = false;
    let disposed = false;
    let lab = false;
    try { lab = win.frameElement?.getAttribute('data-spfx-full-page-preview') === 'true'; } catch { /* Cross-origin hosts expand locally. */ }

    const signal = (value: boolean): void => {
      if (win.parent === win) return;
      win.parent.postMessage({ type: 'spfx-kit:workspace-expanded', expanded: value, appId: 'account-management-spfx' }, win.location.origin);
    };
    const movePreservingScroll = (move: () => void): void => {
      const nodes = [host.appRoot, host.portalHost, ...Array.from(host.appRoot.querySelectorAll('*')), ...Array.from(host.portalHost.querySelectorAll('*'))];
      const positions = nodes.filter(node => node.scrollTop || node.scrollLeft).map(node => ({ node, top: node.scrollTop, left: node.scrollLeft }));
      move();
      positions.forEach(({ node, top, left }) => { node.scrollTop = top; node.scrollLeft = left; });
    };
    const cancelAnimation = (): void => {
      win.cancelAnimationFrame(nextFrame); nextFrame = 0;
      animation?.stop(); animation = undefined;
      state?.frame.removeAttribute('data-workspace-animating');
    };
    const restore = (): void => {
      const current = state;
      if (!current) return;
      current.observer.disconnect();
      current.background.forEach((inert, element) => inert === undefined ? element.removeAttribute('inert') : element.setAttribute('inert', inert));
      current.overflow.forEach(({ node, value, priority }) => value ? node.style.setProperty('overflow', value, priority) : node.style.removeProperty('overflow'));
      current.layer.dispose();
      movePreservingScroll(() => current.placements.forEach(({ node, marker }) => {
        // Never reattach a portal that its owner already disposed.
        if (node.parentNode === current.frame) marker.parentNode?.insertBefore(node, marker);
        marker.remove();
      }));
      current.placeholder.remove(); current.frame.remove();
      win.scrollTo(current.scroll.x, current.scroll.y);
      state = undefined; enteredNative = false;
      if (!disposed) {
        setExpanded(false); setNative(false);
        host.appRoot.querySelector<HTMLElement>('[data-workspace-toggle]')?.focus({ preventScroll: true });
      }
    };
    const settle = (): void => {
      cancelAnimation();
      if (targetExpanded) state?.layer.resume();
      else restore();
    };
    const setBounds = (bounds: Bounds): void => {
      if (!state) return;
      for (const key of ['left', 'top', 'width', 'height'] as const) state.frame.style[key] = `${bounds[key]}px`;
    };
    const transition = (from: Bounds, instant: boolean): void => {
      cancelAnimation();
      const apply = (): void => {
        if (!state) return;
        if (instant || lab || motion.matches || doc.fullscreenElement) { settle(); return; }
        const to: Bounds = targetExpanded ? { left: 0, top: 0, width: win.innerWidth, height: win.innerHeight } : state.placeholder.getBoundingClientRect();
        setBounds(from);
        state.frame.setAttribute('data-workspace-animating', 'true');
        // One scalar spring, matching Atlas/Project/Dashboard. No React renders or layout reads per frame.
        animation = animateValue({
          keyframes: [0, 1], type: spring, visualDuration: 0.7, bounce: 0.04,
          onUpdate: value => setBounds({
            left: from.left + (to.left - from.left) * value,
            top: from.top + (to.top - from.top) * value,
            width: from.width + (to.width - from.width) * value,
            height: from.height + (to.height - from.height) * value
          }),
          onComplete: settle
        });
      };
      // The full-page preview animates its iframe; do not run a second spring inside it.
      if (lab) nextFrame = win.requestAnimationFrame(apply); else apply();
    };
    const expand = (instant = false): void => {
      if (targetExpanded) { if (instant) settle(); return; }
      if (!state && doc.querySelector('[data-account-expanded]')) { setError('Restore the other expanded Account Management workspace first.'); return; }
      const from = (state?.frame || host.appRoot).getBoundingClientRect();
      cancelAnimation(); targetExpanded = true; signal(true);
      if (!state) {
        const focused = doc.activeElement as HTMLElement | null;
        const scroll = { x: win.scrollX, y: win.scrollY };
        const placeholder = doc.createElement('div');
        placeholder.style.height = `${from.height}px`; placeholder.setAttribute('aria-hidden', 'true');
        host.appRoot.before(placeholder);
        const frame = doc.createElement('div');
        frame.className = className; frame.dataset.accountExpanded = host.instanceId;
        frame.setAttribute('role', 'region'); frame.setAttribute('aria-label', 'Expanded Account Management workspace'); frame.tabIndex = -1;
        const placements = [host.appRoot, host.portalHost].map(node => {
          const marker = doc.createComment('Account Management workspace position');
          node.before(marker); return { node, marker };
        });
        doc.body.appendChild(frame);
        movePreservingScroll(() => placements.forEach(({ node }) => frame.appendChild(node)));
        const layer = elevateWorkspace(frame);
        const background = new Map<Element, string | undefined>();
        const isolate = (): void => Array.from(doc.body.children).forEach(child => {
          if (child !== frame && !background.has(child)) { background.set(child, child.getAttribute('inert') ?? undefined); child.setAttribute('inert', ''); }
        });
        isolate();
        const observer = new MutationObserver(isolate); observer.observe(doc.body, { childList: true });
        const overflow = [doc.documentElement, doc.body].map(node => ({ node, value: node.style.getPropertyValue('overflow'), priority: node.style.getPropertyPriority('overflow') }));
        overflow.forEach(({ node }) => node.style.setProperty('overflow', 'hidden'));
        state = { frame, placeholder, placements, layer, background, overflow, observer, scroll };
        (focused && frame.contains(focused) ? focused : frame).focus({ preventScroll: true });
      }
      setExpanded(true); setError(''); transition(from, instant);
    };
    const collapse = (instant = false): void => {
      if (!state) return;
      const from = state.frame.getBoundingClientRect();
      targetExpanded = false; signal(false); transition(from, instant);
    };
    const exit = (): void => {
      if (changingNative) return;
      if (state && doc.fullscreenElement === state.frame) {
        changingNative = true; setPending(true);
        doc.exitFullscreen().then(() => { if (!disposed) collapse(true); }, () => {
          if (!disposed) setError('Use your browser’s full-screen control to exit, then try again.');
        }).then(() => { changingNative = false; if (!disposed) setPending(false); }).catch(() => undefined);
      } else collapse();
    };
    const enterBrowser = (): void => {
      if (changingNative) return;
      expand(true);
      const current = state;
      if (!current) return;
      const unavailable = 'Browser full screen is unavailable here. The workspace still fills the page.';
      if (!doc.fullscreenEnabled || !current.frame.requestFullscreen || doc.fullscreenElement) { setError(unavailable); return; }
      changingNative = true; setPending(true);
      current.layer.suspend();
      current.frame.requestFullscreen().catch(() => {
        if (!disposed) { current.layer.resume(); setError(unavailable); }
      }).then(() => { changingNative = false; if (!disposed) setPending(false); }).catch(() => undefined);
    };
    const keydown = (event: KeyboardEvent): void => {
      const frame = state?.frame;
      if (event.key !== 'Escape' || event.defaultPrevented || !frame || !frame.contains(event.target as Node)) return;
      if (frame.querySelector('[role="menu"], [role="dialog"], [role="alertdialog"], [role="listbox"]')) return;
      event.preventDefault(); exit();
    };
    const focus = (event: FocusEvent): void => {
      const frame = state?.frame;
      if (frame && !frame.contains(event.target as Node)) (host.appRoot.querySelector<HTMLElement>('[data-workspace-toggle]') || frame).focus({ preventScroll: true });
    };
    const changed = (): void => {
      if (state && doc.fullscreenElement === state.frame) { enteredNative = true; setNative(true); }
      else if (enteredNative) collapse(true);
    };
    const message = (event: MessageEvent): void => {
      if (win.parent !== win && event.origin === win.location.origin && event.source === win.parent && event.data?.type === 'spfx-kit:workspace-restore') exit();
    };
    const resized = (): void => { if (!lab) settle(); };
    controls.current = { toggle: () => { if (!changingNative) { if (targetExpanded) exit(); else expand(); } }, enterBrowser };
    doc.addEventListener('keydown', keydown); doc.addEventListener('focusin', focus); doc.addEventListener('fullscreenchange', changed);
    win.addEventListener('resize', resized); win.addEventListener('message', message); motion.addEventListener('change', settle);
    const auto = new URLSearchParams(win.location.search).get('expand');
    if ((auto === '1' || auto === 'true') && !autoExpandedDocuments.has(doc)) { autoExpandedDocuments.add(doc); expand(true); }
    return () => {
      disposed = true; cancelAnimation();
      doc.removeEventListener('keydown', keydown); doc.removeEventListener('focusin', focus); doc.removeEventListener('fullscreenchange', changed);
      win.removeEventListener('resize', resized); win.removeEventListener('message', message); motion.removeEventListener('change', settle);
      if (state && doc.fullscreenElement === state.frame) doc.exitFullscreen().catch(() => undefined);
      if (targetExpanded) signal(false);
      targetExpanded = false; restore(); controls.current = undefined;
    };
  }, [host, className]);
  return { expanded, native, pending, error, toggle: () => controls.current?.toggle(), enterBrowser: () => controls.current?.enterBrowser(), clearError: () => setError('') };
}
