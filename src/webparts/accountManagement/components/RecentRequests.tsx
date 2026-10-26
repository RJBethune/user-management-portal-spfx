import * as React from 'react';
import { Button } from '@spfx-kit/ui-profile/button';
import { Notice } from '@spfx-kit/ui-patterns/notice';
import { StatusBadge, defineStatusTones } from '@spfx-kit/ui-patterns/status-badge';
import { Spinner } from '../../../ui/Controls';
import { AccountManagementService } from '../services/AccountManagementService';
import { IAccountManagementProps } from './IAccountManagementProps';
import { IRequestSummary } from '../models/types';
import { friendlyError } from '../shared/errors';
import styles from './AccountManagement.module.scss';

const tones = defineStatusTones({ Completed: 'success', Failed: 'danger', Error: 'danger', Pending: 'warning' });

export function RecentRequests(props: IAccountManagementProps & { active: boolean }): JSX.Element {
  const [rows, setRows] = React.useState<IRequestSummary[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const [revision, setRevision] = React.useState(0);
  const configKey = JSON.stringify(props.listConfig);
  React.useEffect(() => {
    if (!props.active) return undefined;
    let current = true;
    setLoading(true); setRows([]); setError('');
    new AccountManagementService(props.context, props.listConfig).getRecentRequests(25).then(
      result => { if (current) setRows(result); }, reason => { if (current) setError(friendlyError(reason)); }
    ).then(() => { if (current) setLoading(false); }).catch(() => undefined);
    return () => { current = false; };
  }, [props.active, props.context, configKey, revision]);
  return <section className={styles.accountManagement} aria-label="Recent requests">
    <div className={styles.toolbar}><div className={styles.header}><h2>Recent requests</h2><p>Your latest 25 membership requests across groups.</p></div>
      <Button variant="outline" disabled={loading} onClick={() => setRevision(value => value + 1)}>Refresh requests</Button></div>
    {loading ? <Spinner label="Loading recent requests…" /> : error ? <Notice tone="danger" live="assertive">{error}</Notice> : rows.length ? <ul className={styles.requestHistory}>
      {rows.map(row => {
        const stale = row.status === 'Pending' && !!row.requestedOn && Date.now() - Date.parse(row.requestedOn) > 5 * 60 * 1000;
        const updated = row.modified || row.requestedOn;
        return <li key={row.id} className={styles.recentRow}>
          <StatusBadge status={stale ? 'Error' : row.status} tones={tones} />
          <div className={styles.memberDetails}><strong>{row.action}: {row.memberDisplayName}</strong><span>{row.groupName}</span>
            {row.resultMessage && <span>{row.resultMessage}</span>}
            {stale && <span>No response from the workflow — refresh to confirm.</span>}
            {updated && !isNaN(Date.parse(updated)) && <time dateTime={updated}>{new Date(updated).toLocaleString()}</time>}
          </div>
        </li>;
      })}
    </ul> : <Notice>You have no recent membership requests.</Notice>}
  </section>;
}
