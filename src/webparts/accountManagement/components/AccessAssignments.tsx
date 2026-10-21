import * as React from 'react';
import { useSpfxUiHost, useSpfxUiId } from '@spfx-kit/ui-profile';
import { Button } from '@spfx-kit/ui-profile/button';
import { Input } from '@spfx-kit/ui-profile/input';
import { Notice } from '@spfx-kit/ui-patterns/notice';
import { IAccessOverview, IAssignmentEditor } from '../models/accessAssignments';
import { AssignmentEditor } from './AssignmentEditor';
import { AccessAdministrationService } from '../services/AccessAdministrationService';
import { Link, Spinner } from '../../../ui/Controls';
import { toMessage } from '../shared/errors';
import styles from './AccessAssignments.module.scss';

export function AccessAssignments({ service }: { service: AccessAdministrationService }): JSX.Element {
  const host = useSpfxUiHost();
  const prefix = useSpfxUiId('access-actions');
  const [data, setData] = React.useState<IAccessOverview>();
  const [error, setError] = React.useState<string>();
  const [query, setQuery] = React.useState('');
  const [revision, setRevision] = React.useState(0);
  const [loading, setLoading] = React.useState(true);
  const [editor, setEditor] = React.useState<IAssignmentEditor>();
  const [opening, setOpening] = React.useState(false);
  const [saved, setSaved] = React.useState(false);
  const returnTo = React.useRef('');
  const heading = React.useRef<HTMLHeadingElement>(null);
  const alive = React.useRef(true);
  React.useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, []);
  const openEditor = async (id?: number): Promise<void> => {
    if (opening) return;
    returnTo.current = `${prefix}-${id === undefined ? 'new' : id}`;
    setOpening(true); setError(undefined); setSaved(false);
    try { const next = await service.openEditor(id); if (alive.current) setEditor(next); }
    catch (reason) { if (alive.current) setError(toMessage(reason, 'The assignment could not be opened.')); }
    if (alive.current) setOpening(false);
  };
  const closeEditor = (): void => {
    setEditor(undefined);
    host.targetDocument.defaultView?.requestAnimationFrame(() => host.targetDocument.getElementById(returnTo.current)?.focus());
  };
  React.useEffect(() => {
    let current = true;
    setLoading(true);
    setError(undefined);
    setData(undefined);
    service.load().then((result) => { if (current) setData(result); }, (reason) => {
      if (current) setError(toMessage(reason, 'Access assignments could not be loaded.'));
    }).then(() => { if (current) setLoading(false); }).catch(() => undefined);
    return () => { current = false; };
  }, [service, revision]);
  React.useEffect(() => { if (saved && !loading && heading.current?.offsetParent) heading.current.focus(); }, [saved, loading]);
  const search = query.trim().toLocaleLowerCase();
  const assignments = data ? data.assignments.filter((row) =>
    [row.userName, row.userEmail, ...row.groups.map((group) => `${group.title} ${group.groupId}`)].join(' ').toLocaleLowerCase().includes(search)) : [];
  return <section className={styles.overview} aria-label="Access assignments">
    <div className={styles.heading}>
      <div><h2 ref={heading} tabIndex={-1}>Access assignments</h2><p>See which groups each admin can manage. Add or edit assignments here.</p></div>
      <Button variant="outline" disabled={loading || opening || !!editor} onClick={() => { setSaved(false); setRevision((value) => value + 1); }}>Refresh assignments</Button>
    </div>
    {loading && <Spinner label="Loading access assignments…" />}
    {error && <Notice tone="danger" live="assertive">{error}</Notice>}
    {saved && <Notice tone="success" live="polite">Assignment saved. The affected admin can reload their membership page to use the updated assignments.</Notice>}
    {opening && <Spinner label="Opening assignment…" />}
    {editor && <AssignmentEditor initial={editor} service={service} existingEmails={data?.assignments.map((row) => row.userEmail) || []} onCancel={closeEditor} onSaved={() => {
      setEditor(undefined); setSaved(true); setQuery(''); setRevision((value) => value + 1);
    }} />}
    {data && !editor && <>
      <div className={styles.toolbar}>
        <Input type="search" aria-label="Search admins or assigned groups" placeholder="Search admins or groups" value={query} onChange={(event) => setQuery(event.target.value)} />
        <Button id={`${prefix}-new`} disabled={opening || !data.canAdd} onClick={() => { openEditor().catch(() => undefined); }}>Add an admin assignment</Button>
        <Link href={data.listUrl} target="_blank">Open assignments list</Link>
      </div>
      <p className={styles.note}>All assignments on the configured list site are shown, including groups hidden by this page’s office filter.{!data.canAdd && ' Adding an assignment also requires Add Items permission.'}</p>
      {data.duplicateUserIds.length > 0 && <Notice tone="warning">Some admins have multiple records. Their access combines every assigned group across those records. Review all records before removing access.</Notice>}
      <div className={styles.tableWrap}>
        <table className={styles.table} role="table">
          <caption>{assignments.length} of {data.assignments.length} assignment records</caption>
          <thead role="rowgroup"><tr role="row"><th role="columnheader" scope="col">Admin</th><th role="columnheader" scope="col">Groups this admin can manage</th><th role="columnheader" scope="col">Assignment</th></tr></thead>
          <tbody role="rowgroup">{assignments.map((row) => <tr role="row" key={row.id}>
            <td role="cell"><strong>{row.userName}</strong><span className={styles.secondary}>{row.userEmail}</span>{row.userId && data.duplicateUserIds.indexOf(row.userId) !== -1 && <span className={styles.secondary}>Multiple assignment records</span>}</td>
            <td role="cell"><span className={styles.mobileLabel} aria-hidden="true">Assigned groups</span>{row.groups.length > 0 ? <ul>{row.groups.map((group) => <li key={group.id}><span>{group.title}</span>{group.isTeamsConnected && <span className={styles.secondary}>Teams-connected</span>}<span className={styles.secondary}>{group.groupId || 'Group ID missing'}</span></li>)}</ul> : <span>No groups assigned</span>}
              {row.missingGroupIds.length > 0 && <p>Unavailable group records: {row.missingGroupIds.join(', ')}. Check the assignment and Managed Groups list.</p>}
            </td>
            <td role="cell"><Button id={`${prefix}-${row.id}`} variant="outline" disabled={opening} aria-label={`Edit assignment ${row.id} for ${row.userName}`} onClick={() => { openEditor(row.id).catch(() => undefined); }}>Edit assignment</Button></td>
          </tr>)}</tbody>
        </table>
      </div>
      {!assignments.length && <Notice>{data.assignments.length ? 'No admins or groups match this search.' : 'No admin assignments exist. Add an assignment to select an admin and their groups.'}</Notice>}
      <section className={styles.unassigned} aria-label="Groups without an assigned admin">
        <h3>Groups without an assigned admin ({data.unassigned.length})</h3>
        {data.unassigned.length ? <ul>{data.unassigned.map((group) => <li key={group.id}>{group.title}<span className={styles.secondary}>{group.groupId || 'Group ID missing'}</span></li>)}</ul> : <p>Every managed group appears in an assignment record.</p>}
      </section>
      <Notice title="What an assignment controls">An assignment enables this app’s membership requests for the selected groups. Power Automate rechecks Microsoft 365 requests. SharePoint group changes also require the signed-in user’s SharePoint permissions. This view does not report Teams channel permissions or grant site access.</Notice>
    </>}
  </section>;
}
