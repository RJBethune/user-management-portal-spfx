import * as React from 'react';
import { useSpfxUiHost, useSpfxUiId } from '@spfx-kit/ui-profile';
import { Button } from '@spfx-kit/ui-profile/button';
import { Input } from '@spfx-kit/ui-profile/input';
import { Checkbox } from '@spfx-kit/ui-profile/checkbox';
import { Notice } from '@spfx-kit/ui-patterns/notice';
import { IAssignmentEditor } from '../models/accessAssignments';
import { IUser } from '../models/types';
import { AccessAdministrationService, AssignmentSaveError } from '../services/AccessAdministrationService';
import { toMessage } from '../shared/errors';
import { Spinner } from '../../../ui/Controls';
import { ConfirmDialog } from './ConfirmDialog';
import styles from './AccessAssignments.module.scss';

interface IProps {
  initial: IAssignmentEditor;
  service: AccessAdministrationService;
  existingEmails: string[];
  onCancel: () => void;
  onSaved: () => void;
}
const sameIds = (a: number[], b: number[]): boolean => a.length === b.length && a.every((id) => b.indexOf(id) !== -1);

export function AssignmentEditor({ initial, service, existingEmails, onCancel, onSaved }: IProps): JSX.Element {
  const host = useSpfxUiHost();
  const prefix = useSpfxUiId('assignment-editor');
  const heading = React.useRef<HTMLHeadingElement>(null);
  const groupSearch = React.useRef<HTMLInputElement>(null);
  const inFlight = React.useRef(false);
  const [model, setModel] = React.useState(initial);
  const [groupIds, setGroupIds] = React.useState(initial.groupIds);
  const [groupQuery, setGroupQuery] = React.useState('');
  const [adminQuery, setAdminQuery] = React.useState('');
  const [admin, setAdmin] = React.useState<IUser>();
  const [results, setResults] = React.useState<IUser[]>([]);
  const [searching, setSearching] = React.useState(false);
  const [searchError, setSearchError] = React.useState<string>();
  const [error, setError] = React.useState<string>();
  const [blocked, setBlocked] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [confirm, setConfirm] = React.useState<'cancel' | 'reload'>();
  const dirty = !!admin || !!adminQuery || !sameIds(groupIds, model.groupIds);
  const creating = model.id === undefined;
  React.useEffect(() => { heading.current?.focus(); }, []);
  React.useEffect(() => { if (admin) groupSearch.current?.focus(); }, [admin]);
  React.useEffect(() => {
    if (!dirty && !saving) return undefined;
    const window = host.targetDocument.defaultView;
    const warn = (event: BeforeUnloadEvent): void => { event.preventDefault(); event.returnValue = ''; };
    window?.addEventListener('beforeunload', warn);
    return () => window?.removeEventListener('beforeunload', warn);
  }, [dirty, saving, host]);
  React.useEffect(() => {
    let current = true;
    setResults([]); setSearchError(undefined);
    if (admin || adminQuery.trim().length < 2) { setSearching(false); return undefined; }
    setSearching(true);
    const timer = setTimeout(() => {
      service.searchAdmins(adminQuery).then((users) => { if (current) setResults(users); }, (reason) => {
        if (current) setSearchError(toMessage(reason, 'Directory search failed. Try again.'));
      }).then(() => { if (current) setSearching(false); }).catch(() => undefined);
    }, 300);
    return () => { current = false; clearTimeout(timer); };
  }, [adminQuery, admin, service]);

  const save = async (): Promise<void> => {
    if (inFlight.current || blocked) return;
    inFlight.current = true; setSaving(true); setError(undefined);
    try {
      await service.saveAssignment(model, groupIds, admin);
      onSaved();
    } catch (reason) {
      setError(toMessage(reason, 'The assignment could not be saved.'));
      setBlocked(reason instanceof AssignmentSaveError && (reason.code === 'conflict' || reason.code === 'uncertain'));
      // The synchronous lock prevents a second writer; only this operation releases it.
      // eslint-disable-next-line require-atomic-updates
      inFlight.current = false; setSaving(false);
    }
  };
  const discard = async (): Promise<void> => {
    const action = confirm; setConfirm(undefined);
    if (action !== 'reload') { onCancel(); return; }
    inFlight.current = true; setSaving(true); setError(undefined);
    try {
      const fresh = await service.openEditor(model.id);
      setModel(fresh); setGroupIds(fresh.groupIds); setBlocked(false); setGroupQuery('');
    } catch (reason) { setError(toMessage(reason, 'The assignment could not be reloaded.')); }
    inFlight.current = false; setSaving(false);
  };
  const toggle = (id: number, checked: boolean): void => {
    setGroupIds((current) => checked ? (model.multiple ? Array.from(new Set([...current, id])) : [id]) : current.filter((value) => value !== id));
  };
  const query = groupQuery.trim().toLocaleLowerCase();
  const available = model.groups.filter((group) => `${group.title} ${group.groupId}`.toLocaleLowerCase().includes(query));
  const missing = groupIds.filter((id) => !model.groups.some((group) => group.id === id));
  const added = groupIds.filter((id) => model.groupIds.indexOf(id) === -1).length;
  const removed = model.groupIds.filter((id) => groupIds.indexOf(id) === -1).length;
  const duplicate = !!admin && existingEmails.some((email) => email && [admin.mail, admin.userPrincipalName].some((value) => value?.toLowerCase() === email.toLowerCase()));

  return <form className={styles.editor} aria-label={creating ? 'Add admin assignment' : 'Edit admin assignment'} onSubmit={(event) => { event.preventDefault(); save().catch(() => undefined); }}>
    <div><h3 ref={heading} tabIndex={-1}>{creating ? 'Add an admin assignment' : `Edit groups for ${model.userName}`}</h3>
      <p>Choose the groups this admin can manage. Changes take effect when you save.</p></div>
    {error && <Notice tone="danger" live="assertive">{error}</Notice>}
    <fieldset disabled={saving} className={styles.editorFields}>
      <legend>Admin</legend>
      {!creating ? <p><strong>{model.userName}</strong><span className={styles.secondary}>{model.userEmail}</span></p> : <>
        {admin ? <div className={styles.toolbar}><p><strong>{admin.displayName}</strong><span className={styles.secondary}>{admin.userPrincipalName}</span></p>
          <Button type="button" variant="outline" onClick={() => { setAdmin(undefined); setAdminQuery(''); }}>Change admin</Button></div> : <>
          <label htmlFor={`${prefix}-admin`}>Find an admin by name or email</label>
          <Input id={`${prefix}-admin`} type="search" autoComplete="off" value={adminQuery} onChange={(event) => setAdminQuery(event.target.value)} aria-describedby={`${prefix}-search-help`} />
          <p id={`${prefix}-search-help`} className={styles.note}>Type at least 2 characters, then choose a person. Up to 25 results are shown.</p>
          {searching && <Spinner label="Searching directory…" />}
          {searchError && <Notice tone="danger" live="assertive">{searchError}</Notice>}
          {!searching && !searchError && adminQuery.trim().length >= 2 && !results.length && <p role="status">No admins found. Try a different name or email.</p>}
          {results.length > 0 && <ul className={styles.people} aria-label="Directory results">{results.map((person) => <li key={person.id}>
            <Button type="button" variant="outline" aria-label={`Choose ${person.displayName}, ${person.userPrincipalName}`} onClick={() => setAdmin(person)}>{person.displayName} · {person.userPrincipalName}</Button>
          </li>)}</ul>}
        </>}
        {duplicate && <Notice tone="warning">This admin already has an assignment. Access combines all their records. Edit the existing record when possible.</Notice>}
      </>}
    </fieldset>
    <fieldset disabled={saving} className={styles.editorFields}>
      <legend>Groups this admin can manage</legend>
      <p className={styles.note}>{model.multiple ? 'Select every group this admin needs.' : 'This list allows one group per assignment.'} {model.required ? 'At least one group is required.' : 'Clear all groups to remove access from this record.'}</p>
      <Input ref={groupSearch} type="search" aria-label="Find groups to assign" placeholder="Find a group" value={groupQuery} onChange={(event) => setGroupQuery(event.target.value)} />
      <p role="status" className={styles.selectionSummary}>{groupIds.length} selected · {available.length} of {model.groups.length} groups shown</p>
      <ul className={styles.groupChoices}>{available.map((group) => <li key={group.id}>
        <Checkbox id={`${prefix}-group-${group.id}`} disabled={saving} checked={groupIds.indexOf(group.id) !== -1} onCheckedChange={(checked) => toggle(group.id, checked)} />
        <label htmlFor={`${prefix}-group-${group.id}`}><span>{group.title}</span>{group.isTeamsConnected && <span className={styles.secondary}>Teams-connected</span>}<span className={styles.secondary}>{group.groupId || 'Group ID missing'}</span></label>
      </li>)}</ul>
      {!available.length && <p>{model.groups.length ? 'No groups match this search. Existing selections are kept.' : 'No managed groups are available. Ask a list owner to check the Managed Groups list.'}</p>}
      {missing.length > 0 && <Notice tone="warning">Unavailable group records are still selected: {missing.join(', ')}. Remove them before saving.
        <Button type="button" variant="outline" onClick={() => setGroupIds((ids) => ids.filter((id) => missing.indexOf(id) === -1))}>Remove unavailable groups</Button>
      </Notice>}
    </fieldset>
    {!creating && <Notice>{added} {added === 1 ? 'group' : 'groups'} added · {removed} {removed === 1 ? 'group' : 'groups'} removed. Other assignment records for this admin remain in effect.</Notice>}
    <div className={styles.toolbar}>
      <Button type="submit" disabled={saving || blocked || missing.length > 0 || (creating ? !admin : sameIds(groupIds, model.groupIds)) || (model.required && !groupIds.length)}>{saving ? 'Saving…' : 'Save assignment'}</Button>
      <Button type="button" variant="outline" disabled={saving} onClick={() => dirty ? setConfirm('cancel') : onCancel()}>Cancel</Button>
      {blocked && !creating && <Button type="button" variant="outline" disabled={saving} onClick={() => setConfirm('reload')}>Reload assignment</Button>}
    </div>
    <ConfirmDialog open={!!confirm} ariaLabel="Discard assignment changes?" onDismiss={() => setConfirm(undefined)}>
      <h3>Discard assignment changes?</h3><p>{confirm === 'reload' ? 'Reloading replaces your draft with the current saved assignment.' : 'Your changes have not been saved.'}</p>
      <div className={styles.toolbar}><Button type="button" variant="outline" onClick={() => setConfirm(undefined)}>Keep editing</Button><Button type="button" onClick={() => { discard().catch(() => undefined); }}>{confirm === 'reload' ? 'Discard and reload' : 'Discard changes'}</Button></div>
    </ConfirmDialog>
  </form>;
}
