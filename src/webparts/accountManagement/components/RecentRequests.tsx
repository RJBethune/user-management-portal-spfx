import * as React from 'react';
import { useSpfxUiHost } from '@spfx-kit/ui-profile';
import { Button } from '@spfx-kit/ui-profile/button';
import { Notice } from '@spfx-kit/ui-patterns/notice';
import { StatusBadge, defineStatusTones } from '@spfx-kit/ui-patterns/status-badge';
import { Spinner } from '../../../ui/Controls';
import { AccountManagementService } from '../services/AccountManagementService';
import { IAccountManagementProps } from './IAccountManagementProps';
import { IRequestSummary } from '../models/types';
import { friendlyError } from '../shared/errors';
import { delayedRequestText, isDelayedRequest } from '../shared/requestStatus';
import styles from './AccountManagement.module.scss';

const tones = defineStatusTones({ Completed: 'success', Failed: 'danger', Error: 'danger', Pending: 'warning' });

export function RecentRequests(props: IAccountManagementProps & { active: boolean }): JSX.Element {
  const host = useSpfxUiHost();
  const configKey = JSON.stringify(props.listConfig);
  const service = React.useMemo(() => new AccountManagementService(props.context, props.listConfig), [props.context, configKey]);
  const [snapshot, setSnapshot] = React.useState<{ source: AccountManagementService; rows: IRequestSummary[]; checkedAt: number }>();
  const rows = snapshot?.source === service ? snapshot.rows : undefined;
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const [revision, setRevision] = React.useState(0);
  React.useEffect(() => {
    if (!props.active) return undefined;
    let current = true;
    let inFlight = false;
    let pending = true;
    let failed = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const document = host.targetDocument;
    const isVisible = (): boolean => document.visibilityState !== 'hidden';
    const refresh = async (): Promise<void> => {
      if (!current || inFlight || !isVisible()) return;
      inFlight = true; setLoading(true); setError('');
      try {
        const result = await service.getRecentRequests(25);
        if (current) {
          pending = result.some(row => row.status === 'Pending');
          setSnapshot({ source: service, rows: result, checkedAt: Date.now() });
        }
      } catch (reason) {
        if (current) { failed = true; setError(friendlyError(reason)); }
      } finally {
        // The synchronous guard permits only this read to release the lock.
        // eslint-disable-next-line require-atomic-updates
        inFlight = false;
        if (current) {
          setLoading(false);
          // Never overlap reads, poll a hidden view, or repeatedly retry a failing service.
          if (pending && !failed && isVisible()) timer = setTimeout(() => { refresh().catch(() => undefined); }, 15000);
        }
      }
    };
    const visibilityChanged = (): void => {
      clearTimeout(timer);
      if (isVisible() && pending && !failed) refresh().catch(() => undefined);
    };
    setError('');
    document.addEventListener('visibilitychange', visibilityChanged);
    refresh().catch(() => undefined);
    return () => { current = false; clearTimeout(timer); document.removeEventListener('visibilitychange', visibilityChanged); };
  }, [props.active, service, revision, host]);
  return <section className={styles.accountManagement} aria-label="Recent requests">
    <div className={styles.toolbar}><div className={styles.header}><h2>Recent requests</h2><p>Your latest 25 membership requests across groups.</p></div>
      <Button variant="outline" disabled={loading} onClick={() => setRevision(value => value + 1)}>Refresh requests</Button></div>
    {error && <Notice tone="danger" live="assertive">{rows ? `Could not refresh requests. Showing the last loaded results. ${error}` : error}</Notice>}
    {rows && <p className={styles.requestRefreshState}>{loading ? 'Refreshing requests…' : `Last checked ${new Date(snapshot!.checkedAt).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit', second: '2-digit' })}.`}
      {!error && rows.some(row => row.status === 'Pending') && ' Pending requests refresh every 15 seconds while this view is visible.'}</p>}
    {!rows && loading ? <Spinner label="Loading recent requests…" /> : rows?.length ? <ul className={styles.requestHistory}>
      {rows.map(row => {
        const delayed = isDelayedRequest(row.status, row.requestedOn);
        const updated = row.modified || row.requestedOn;
        return <li key={row.id} className={styles.recentRow}>
          <StatusBadge className={styles.requestStatus} status={row.status} tones={tones} />
          <div className={styles.requestDetails}><strong>{row.action}: {row.memberDisplayName}</strong><span>{row.groupName}</span>
            {row.resultMessage && <span>{row.resultMessage}</span>}
            {delayed && <span>{delayedRequestText}</span>}
          </div>
          {updated && !isNaN(Date.parse(updated)) && <time className={styles.requestTime} dateTime={updated}>
            {new Date(updated).toLocaleString(undefined, { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' })}
          </time>}
        </li>;
      })}
    </ul> : rows && !error ? <Notice>You have no recent membership requests.</Notice> : undefined}
  </section>;
}
