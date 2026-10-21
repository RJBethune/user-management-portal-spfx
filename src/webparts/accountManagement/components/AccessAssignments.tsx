import * as React from 'react';
import { useSpfxUiHost, useSpfxUiId } from '@spfx-kit/ui-profile';
import { Button } from '@spfx-kit/ui-profile/button';
import { Input } from '@spfx-kit/ui-profile/input';
import { Notice } from '@spfx-kit/ui-patterns/notice';
import { ArrowLeft20Regular, Person20Regular } from '@fluentui/react-icons';
import { IAccessOverview, IAssignmentEditor } from '../models/accessAssignments';
import { AssignmentEditor } from './AssignmentEditor';
import { AccessAdministrationService } from '../services/AccessAdministrationService';
import { Link, Spinner } from '../../../ui/Controls';
import { toMessage } from '../shared/errors';
import styles from './AccessAssignments.module.scss';

interface IEditorSession { key: string; model: IAssignmentEditor; dirty: boolean; busy: boolean }

export function AccessAssignments({ service }: { service: AccessAdministrationService }): JSX.Element {
  const host = useSpfxUiHost();
  const prefix = useSpfxUiId('access-actions');
  const [data, setData] = React.useState<IAccessOverview>();
  const [error, setError] = React.useState<string>();
  const [query, setQuery] = React.useState('');
  const [revision, setRevision] = React.useState(0);
  const [loading, setLoading] = React.useState(true);
  const [sessions, setSessions] = React.useState<IEditorSession[]>([]);
  const [selected, setSelected] = React.useState<string>();
  const [browsing, setBrowsing] = React.useState(true);
  const [activation, setActivation] = React.useState(0);
  const [opening, setOpening] = React.useState(false);
  const [saved, setSaved] = React.useState(false);
  const heading = React.useRef<HTMLHeadingElement>(null);
  const searchInput = React.useRef<HTMLInputElement>(null);
  const alive = React.useRef(true);
  const openLock = React.useRef(false);
  const busy = sessions.some((session) => session.busy);
  const dirty = sessions.some((session) => session.dirty);
  const hasNewSession = sessions.some((session) => session.key === 'new');
  React.useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, []);
  const focusList = (key?: string): void => {
    host.targetDocument.defaultView?.requestAnimationFrame(() => {
      const row = host.targetDocument.getElementById(`${prefix}-${key}`);
      if (row?.offsetParent) row.focus(); else searchInput.current?.focus();
    });
  };
  const openEditor = async (id?: number): Promise<void> => {
    if (openLock.current || busy) return;
    const key = id === undefined ? 'new' : String(id);
    setSaved(false); setError(undefined);
    if (sessions.some((session) => session.key === key)) { setSelected(key); setBrowsing(false); setActivation((value) => value + 1); return; }
    openLock.current = true; setOpening(true);
    try {
      const model = await service.openEditor(id);
      if (alive.current) {
        setSessions((current) => [...current, { key, model, dirty: false, busy: false }]);
        setSelected(key); setBrowsing(false);
      }
    } catch (reason) { if (alive.current) setError(toMessage(reason, 'The assignment could not be opened.')); }
    // This lock serializes editor reads, so an older response cannot select the wrong admin.
    // eslint-disable-next-line require-atomic-updates
    openLock.current = false;
    if (alive.current) setOpening(false);
  };
  const closeEditor = (key: string): void => {
    setSessions((current) => current.filter((session) => session.key !== key));
    setSelected(undefined); setBrowsing(true); focusList(key);
  };
  const editorStateChanged = React.useCallback((key: string, nextDirty: boolean, nextBusy: boolean): void => {
    setSessions((current) => current.map((session) => session.key === key && (session.dirty !== nextDirty || session.busy !== nextBusy)
      ? { ...session, dirty: nextDirty, busy: nextBusy } : session));
  }, []);
  React.useEffect(() => {
    let current = true;
    setLoading(true); setError(undefined);
    service.load().then((result) => { if (current) setData(result); }, (reason) => {
      if (current) setError(toMessage(reason, 'Access assignments could not be loaded.'));
    }).then(() => { if (current) setLoading(false); }).catch(() => undefined);
    return () => { current = false; };
  }, [service, revision]);
  React.useEffect(() => { if (saved && !loading && heading.current?.offsetParent) heading.current.focus(); }, [saved, loading]);
  const search = query.trim().toLocaleLowerCase();
  // A concurrent deletion must not strand a draft after saving a different record refreshes the list.
  const retained = sessions.filter(({ model }) => model.id !== undefined && !data?.assignments.some((row) => row.id === model.id)).map(({ model }) => ({
    id: model.id!, userId: model.userId, userName: model.userName, userEmail: model.userEmail,
    groups: model.groups.filter((group) => model.groupIds.indexOf(group.id) !== -1),
    missingGroupIds: model.groupIds.filter((id) => !model.groups.some((group) => group.id === id))
  }));
  const inventory = [...(data?.assignments || []), ...retained];
  const assignments = inventory.filter((row) =>
    [row.userName, row.userEmail, ...row.groups.map((group) => `${group.title} ${group.groupId}`)].join(' ').toLocaleLowerCase().includes(search));
  return <section className={styles.overview} aria-label="Access assignments">
    <div className={styles.heading}>
      <div><h2 ref={heading} tabIndex={-1}>Access assignments</h2><p>Select an admin to review and change the groups they can manage.</p></div>
      <Button variant="outline" disabled={loading || opening || busy || dirty} title={dirty ? 'Save or discard your drafts before refreshing' : undefined} onClick={() => {
        setSaved(false); setSessions([]); setSelected(undefined); setBrowsing(true); setRevision((value) => value + 1);
      }}>Refresh assignments</Button>
    </div>
    {loading && <Spinner label="Loading access assignments…" />}
    {error && <Notice tone="danger" live="assertive">{error}</Notice>}
    {saved && <Notice tone="success" live="polite">Assignment saved. The affected admin can reload their membership page to use the updated assignments.</Notice>}
    {opening && <Spinner label="Opening assignment…" />}
    {data && <>
      <div className={`${styles.workspace} ${browsing || !selected ? styles.showList : styles.showDetails}`}>
        <div className={styles.adminPane}>
          <div className={styles.listToolbar}>
            <div className={styles.listHeading}><h3>Admins</h3><span>{assignments.length} of {inventory.length} records</span></div>
            <Input ref={searchInput} type="search" aria-label="Search admins or assigned groups" placeholder="Search admins or groups" value={query} onChange={(event) => setQuery(event.target.value)} />
            <Button id={`${prefix}-new`} disabled={opening || busy || loading || (!data.canAdd && !hasNewSession)} aria-current={selected === 'new' ? 'true' : undefined} onClick={() => { openEditor().catch(() => undefined); }}>
              {sessions.some((session) => session.key === 'new' && session.dirty) ? 'Continue new assignment' : 'Add an admin assignment'}
            </Button>
            {!data.canAdd && <p className={styles.note}>Adding an assignment also requires Add Items permission.</p>}
          </div>
          <nav className={styles.adminNavigation} aria-label="Admin assignments">
            <ul>{assignments.map((row) => {
              const key = String(row.id);
              const duplicate = !!row.userId && data.duplicateUserIds.indexOf(row.userId) !== -1;
              const unsaved = sessions.some((session) => session.key === key && session.dirty);
              return <li key={row.id}><Button id={`${prefix}-${key}`} className={styles.adminChoice} variant={selected === key ? 'secondary' : 'ghost'}
                aria-current={selected === key ? 'true' : undefined} disabled={opening || busy || loading} aria-label={`Edit assignment ${row.id} for ${row.userName}`} aria-describedby={unsaved ? `${prefix}-draft-${key}` : undefined}
                onClick={() => { openEditor(row.id).catch(() => undefined); }}>
                <span className={styles.choiceContent}><Person20Regular aria-hidden="true" /><span className={styles.choiceText}>
                  <span className={styles.choiceName}>{row.userName}</span><span className={styles.secondary}>{row.userEmail}</span>
                  <span className={styles.secondary}>{row.groups.length} {row.groups.length === 1 ? 'group' : 'groups'}{duplicate ? ` · Record ${row.id}` : ''}{row.missingGroupIds.length ? ` · ${row.missingGroupIds.length} unavailable` : ''}</span>
                  {retained.some((record) => record.id === row.id) && <span className={styles.secondary}>Not in the latest list</span>}
                  {unsaved && <span id={`${prefix}-draft-${key}`} className={styles.draftLabel}>Unsaved changes</span>}
                </span></span>
              </Button></li>;
            })}</ul>
            {!assignments.length && <p className={styles.listEmpty}>{data.assignments.length ? 'No admins or groups match this search.' : 'No admin assignments exist. Add an assignment to select an admin and their groups.'}</p>}
          </nav>
          <div className={styles.listFooter}><Link href={data.listUrl} target="_blank">Open assignments list</Link>
            <p className={styles.note}>All assignment records are shown, including groups hidden by this page’s office filter.</p>
          </div>
        </div>
        <div className={styles.detailsPane}>
          {selected && <div className={styles.backToAdmins}><Button variant="ghost" disabled={busy || opening} onClick={() => { setBrowsing(true); focusList(selected); }}><ArrowLeft20Regular aria-hidden="true" />Back to admins</Button></div>}
          {!selected && <div className={styles.emptyDetail}><h3>Select an admin</h3><p>Review their groups here, or add an admin assignment. Drafts stay available when you switch admins.</p></div>}
          {sessions.map((session) => <div key={session.key} hidden={selected !== session.key} className={styles.editorPane}>
            {session.model.userId && data.duplicateUserIds.indexOf(session.model.userId) !== -1 && <Notice tone="warning">This admin has multiple records. You are editing record {session.model.id}. Access combines groups from every record.</Notice>}
            <AssignmentEditor initial={session.model} sessionKey={session.key} active={selected === session.key} activation={activation} onStateChange={editorStateChanged}
              service={service} existingEmails={data.assignments.map((row) => row.userEmail)} onCancel={() => closeEditor(session.key)} onSaved={() => {
                closeEditor(session.key); setSaved(true); setQuery(''); setRevision((value) => value + 1);
              }} />
          </div>)}
        </div>
      </div>
      <details className={styles.unassigned}><summary><h3>Groups without an assigned admin ({data.unassigned.length})</h3></summary>
        {data.unassigned.length ? <ul>{data.unassigned.map((group) => <li key={group.id}>{group.title}<span className={styles.secondary}>{group.groupId || 'Group ID missing'}</span></li>)}</ul> : <p>Every managed group appears in an assignment record.</p>}
      </details>
      <p className={styles.note}>Assignments enable membership requests for selected groups. Existing SharePoint permissions and Power Automate checks still apply. Assignments do not grant site or Teams channel access.</p>
    </>}
  </section>;
}
