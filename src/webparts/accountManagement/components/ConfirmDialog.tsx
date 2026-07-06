import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './AccountManagement.module.scss';

export interface IConfirmDialogProps {
  open: boolean;
  ariaLabel: string;
  onDismiss: () => void;
  children?: React.ReactNode;
}

/**
 * Minimal accessible modal confirmation. role=alertdialog + aria-modal; focus moves into the
 * panel on open and returns to the trigger on close; Tab is trapped; Escape and overlay-click
 * dismiss. Hand-rolled (React portal + plain DOM focus) on purpose — the @fluentui/react Dialog's
 * tabster focus-trap fails to initialize in this SPFx bundle ("Cannot read properties of undefined
 * (reading 'set')"), so we avoid it entirely.
 */
export const ConfirmDialog: React.FunctionComponent<IConfirmDialogProps> = (props: IConfirmDialogProps) => {
  const panelRef: React.RefObject<HTMLDivElement> = React.useRef<HTMLDivElement>(null);
  const returnFocusRef: React.MutableRefObject<HTMLElement | undefined> = React.useRef<HTMLElement | undefined>(undefined);

  React.useEffect(() => {
    if (!props.open) {
      return undefined;
    }
    returnFocusRef.current = (document.activeElement as HTMLElement) || undefined;
    if (panelRef.current) {
      panelRef.current.focus();
    }
    return () => {
      if (returnFocusRef.current && typeof returnFocusRef.current.focus === 'function') {
        returnFocusRef.current.focus();
      }
    };
  }, [props.open]);

  if (!props.open) {
    return null;
  }

  const onKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      props.onDismiss();
      return;
    }
    if (e.key !== 'Tab' || !panelRef.current) {
      return;
    }
    const focusables: HTMLElement[] = Array.prototype.slice.call(
      panelRef.current.querySelectorAll(
        'button:not([disabled]), [href], input:not([disabled]), textarea:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])'
      )
    ) as HTMLElement[];
    if (!focusables.length) {
      return;
    }
    const first: HTMLElement = focusables[0];
    const last: HTMLElement = focusables[focusables.length - 1];
    const active: Element | null = document.activeElement;
    if (e.shiftKey && (active === first || active === panelRef.current)) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && active === last) {
      e.preventDefault();
      first.focus();
    }
  };

  return ReactDOM.createPortal(
    <div
      className={styles.modalOverlay}
      onMouseDown={(e: React.MouseEvent<HTMLDivElement>) => {
        if (e.target === e.currentTarget) {
          props.onDismiss();
        }
      }}
    >
      <div
        ref={panelRef}
        className={styles.modalPanel}
        role="alertdialog"
        aria-modal="true"
        aria-label={props.ariaLabel}
        tabIndex={-1}
        onKeyDown={onKeyDown}
      >
        {props.children}
      </div>
    </div>,
    document.body
  );
};
