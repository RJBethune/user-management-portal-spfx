import * as React from 'react';
import { Button as KitButton } from '@spfx-kit/ui-profile/button';
import { Input } from '@spfx-kit/ui-profile/input';
import { Textarea as KitTextarea } from '@spfx-kit/ui-profile/textarea';
import { Spinner as KitSpinner } from '@spfx-kit/ui-profile/spinner';
import { Switch as KitSwitch } from '@spfx-kit/ui-profile/switch';
import { Notice } from '@spfx-kit/ui-patterns/notice';
import { Field as KitField, FieldLabel } from '@spfx-kit/ui-profile/field';
import styles from './Controls.module.scss';

/** Translate the existing app's event contracts to Kit controls without changing membership behavior. */
type ButtonProps = React.ComponentProps<typeof KitButton> & {
  appearance?: 'primary' | 'secondary' | 'transparent' | 'destructive';
  icon?: React.ReactNode;
};
export function Button({ appearance, icon, children, ...props }: ButtonProps): JSX.Element {
  const variant = appearance === 'primary' ? 'default' : appearance === 'transparent' ? 'ghost' : appearance === 'destructive' ? 'destructive' : 'outline';
  return <KitButton variant={variant} size={!children ? 'icon' : 'default'} {...props}>{icon && <span aria-hidden="true">{icon}</span>}{children}</KitButton>;
}

interface SearchProps extends Omit<React.InputHTMLAttributes<HTMLInputElement>, 'onChange' | 'size'> {
  onChange: (event: React.ChangeEvent<HTMLInputElement>, data: { value: string }) => void;
}
export function SearchBox({ onChange, className, ...props }: SearchProps): JSX.Element {
  return <div className={className}><Input type="search" {...props} onChange={(event) => onChange(event, { value: event.target.value })} /></div>;
}

interface TextareaProps extends Omit<React.TextareaHTMLAttributes<HTMLTextAreaElement>, 'onChange'> {
  onChange: (event: React.ChangeEvent<HTMLTextAreaElement>, data: { value: string }) => void;
}
export function Textarea({ onChange, ...props }: TextareaProps): JSX.Element {
  return <KitTextarea {...props} onChange={(event) => onChange(event, { value: event.target.value })} />;
}

export function Field({ id, label, required, children }: { id: string; label: string; required?: boolean; children: React.ReactElement }): JSX.Element {
  return <KitField><FieldLabel htmlFor={id}>{label}{required ? ' (required)' : ''}</FieldLabel>{React.cloneElement(children, { id, required })}</KitField>;
}

export function Switch({ checked, label, disabled, onChange }: {
  checked: boolean; label: string; disabled?: boolean;
  onChange: (event: undefined, data: { checked: boolean }) => void;
}): JSX.Element {
  return <label className={styles.switch}><KitSwitch checked={checked} disabled={disabled} onCheckedChange={(value) => onChange(undefined, { checked: value })} /><span>{label}</span></label>;
}

export function Spinner({ label }: { label?: string; size?: 'small' | 'medium' }): JSX.Element {
  return <span className={styles.loading} role="status" aria-label={label || 'Loading'}><KitSpinner aria-hidden="true" />{label && <span>{label}</span>}</span>;
}

export function Link(props: React.AnchorHTMLAttributes<HTMLAnchorElement>): JSX.Element {
  return <a {...props} className={styles.link} rel={props.target === '_blank' ? 'noopener noreferrer' : props.rel} />;
}

export function MessageBar({ intent = 'info', politeness, children }: {
  intent?: 'info' | 'success' | 'warning' | 'error'; politeness?: 'polite' | 'assertive'; layout?: string; children?: React.ReactNode;
}): JSX.Element {
  return <Notice tone={intent === 'error' ? 'danger' : intent} live={politeness || 'off'}>{children}</Notice>;
}
export function MessageBarBody({ children }: { children?: React.ReactNode }): JSX.Element { return <div>{children}</div>; }
export function MessageBarTitle({ children }: { children?: React.ReactNode }): JSX.Element { return <strong>{children}</strong>; }
export function MessageBarActions({ containerAction }: { containerAction: React.ReactNode }): JSX.Element { return <div className={styles.actions}>{containerAction}</div>; }
