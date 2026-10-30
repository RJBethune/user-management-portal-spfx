import * as React from 'react';
import { useSpfxUiHost, useSpfxUiId } from '@spfx-kit/ui-profile';
import { Button } from '@spfx-kit/ui-profile/button';
import { Input } from '@spfx-kit/ui-profile/input';
import { ArrowLeft20Regular, PeopleTeam20Regular } from '@fluentui/react-icons';
import { groupAssignmentCoverage, IAccessOverview } from '../models/accessAssignments';
import { Link } from '../../../ui/Controls';
import styles from './AccessAssignments.module.scss';

interface IProps {
  overview: IAccessOverview;
  active: boolean;
  disabled: boolean;
  hasDrafts: boolean;
  onOpenAssignment: (id: number) => void;
  onBrowseAdmins: () => void;
}

/** A second way into the existing assignment editor, using the same saved inventory. */
export function AccessGroups({ overview, active, disabled, hasDrafts, onOpenAssignment, onBrowseAdmins }: IProps): JSX.Element {
  const host = useSpfxUiHost();
  const prefix = useSpfxUiId('access-groups');
  const [query, setQuery] = React.useState('');
  const [filter, setFilter] = React.useState<'all' | 'unassigned' | 'duplicates'>('all');
  const [selected, setSelected] = React.useState<number>();
  const [browsing, setBrowsing] = React.useState(true);
  const heading = React.useRef<HTMLHeadingElement>(null);
  const searchInput = React.useRef<HTMLInputElement>(null);
  const coverage = React.useMemo(() => groupAssignmentCoverage(overview.assignments, overview.groups), [overview]);
  const search = query.trim().toLocaleLowerCase();
  const shown = coverage.filter(row =>
    (filter === 'all' || (filter === 'unassigned' ? row.adminCount === 0 : row.duplicateGrants)) &&
    [row.group.title, row.group.groupId, ...row.assignments.map(admin => `${admin.userName} ${admin.userEmail}`)].join(' ').toLocaleLowerCase().includes(search));
  const current = coverage.find(row => row.group.id === selected);
  React.useEffect(() => { if (active && current && !browsing) heading.current?.focus(); }, [active, selected, browsing]);
  const back = (): void => {
    setBrowsing(true);
    host.targetDocument.defaultView?.requestAnimationFrame(() => {
      const row = host.targetDocument.getElementById(`${prefix}-${selected}`);
      if (row?.offsetParent) row.focus(); else searchInput.current?.focus();
    });
  };
  return <div className={`${styles.workspace} ${browsing || !current ? styles.showList : styles.showDetails}`}>
    <div className={styles.adminPane}>
      <div className={styles.listToolbar}>
        <div className={styles.listHeading}><h3>Groups</h3><span>{shown.length} of {coverage.length}</span></div>
        <Input ref={searchInput} type="search" aria-label="Search groups or assigned admins" placeholder="Search groups or admins" value={query} onChange={event => setQuery(event.target.value)} />
        <div role="group" aria-label="Filter group assignments" className={styles.coverageFilters}>
          {(['all', 'unassigned', 'duplicates'] as const).map(value => <Button key={value} variant={filter === value ? 'secondary' : 'ghost'} size="sm" aria-pressed={filter === value} onClick={() => setFilter(value)}>
            {value === 'all' ? 'All groups' : value === 'unassigned' ? 'No assigned admin' : 'Duplicate assignments'}
          </Button>)}
        </div>
      </div>
      <nav className={styles.adminNavigation} aria-label="Group assignments">
        <ul>{shown.map(row => <li key={row.group.id}><Button id={`${prefix}-${row.group.id}`} className={styles.adminChoice} variant={selected === row.group.id ? 'secondary' : 'ghost'}
          aria-current={selected === row.group.id ? 'true' : undefined} aria-label={`Review admins for ${row.group.title}`} disabled={disabled}
          onClick={() => { setSelected(row.group.id); setBrowsing(false); }}>
          <span className={styles.choiceContent}><PeopleTeam20Regular aria-hidden="true" /><span className={styles.choiceText}>
            <span className={styles.choiceName}>{row.group.title}</span>
            <span className={styles.secondary}>{row.adminCount ? `${row.adminCount} assigned ${row.adminCount === 1 ? 'admin' : 'admins'}` : 'No assigned admin'}</span>
            {row.duplicateGrants && <span className={styles.secondary}>Duplicate assignments</span>}
          </span></span>
        </Button></li>)}</ul>
        {!shown.length && <div className={styles.listEmpty}><p>{coverage.length ? 'No groups match these filters.' : 'No managed groups are available.'}</p>
          {coverage.length > 0 && <Button variant="ghost" onClick={() => { setQuery(''); setFilter('all'); }}>Clear group filters</Button>}
        </div>}
      </nav>
      <div className={styles.listFooter}><p className={styles.note}>Counts show distinct admins in saved assignment records.</p></div>
    </div>
    <div className={styles.detailsPane}>
      {current ? <>
        <div className={styles.backToAdmins}><Button variant="ghost" disabled={disabled} onClick={back}><ArrowLeft20Regular aria-hidden="true" />Back to groups</Button></div>
        <section className={styles.groupCoverage} aria-label="Assigned admins for selected group">
          <div><h3 ref={heading} tabIndex={-1}>{current.group.title}</h3>
            <p>{current.adminCount ? `${current.adminCount} ${current.adminCount === 1 ? 'admin can' : 'admins can'} manage this group through saved assignments.` : 'No admin is assigned to manage this group.'}</p>
            {current.group.isTeamsConnected && <p className={styles.note}>Teams-connected group</p>}
          </div>
          {hasDrafts && <p className={styles.note}>Unsaved changes are kept in Admins. This view shows saved assignments.</p>}
          {!shown.some(row => row.group.id === selected) && <p className={styles.note}>This selected group is outside the current filters.</p>}
          {current.duplicateGrants && <p className={styles.note}>An admin has this group in more than one record. To remove their access, update every record granting it.</p>}
          {current.assignments.length > 0 && <ul className={styles.coverageRecords} aria-label="Saved assignment records">{current.assignments.map(row => <li key={row.id}>
            <div><strong>{row.userName}</strong><span className={styles.secondary}>{row.userEmail}</span><span className={styles.secondary}>Assignment record {row.id}</span>
              {!row.userId && <span className={styles.secondary}>This record has no admin. Repair it in the assignments list.</span>}
            </div>
            {row.userId ? <Button variant="outline" size="sm" disabled={disabled} aria-label={`Edit assignment ${row.id} for ${row.userName}`} onClick={() => onOpenAssignment(row.id)}>Edit groups</Button>
              : <Link href={overview.listUrl} target="_blank">Open assignments list</Link>}
          </li>)}</ul>}
          <div><Button variant="outline" disabled={disabled} onClick={onBrowseAdmins}>Browse admins to assign this group</Button>
            <p className={styles.note}>Choose an admin, then select this group in their assignment.</p></div>
        </section>
      </> : <div className={styles.emptyDetail}><h3>Select a group</h3><p>See who can manage it, then open an assignment to change that admin’s groups.</p></div>}
    </div>
  </div>;
}
