import * as React from 'react';
import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  SearchBox,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Icon
} from '@fluentui/react';
import { LivePersona } from '@pnp/spfx-controls-react/lib/LivePersona';
import styles from './AccountManagement.module.scss';
import { IAccountManagementProps } from './IAccountManagementProps';
import { GraphService } from '../services/GraphService';
import { AccountManagementService } from '../services/AccountManagementService';
import { IOfficeGroup, IUser, IRequestSummary, MembershipAction } from '../models/types';
import { isSharePointGroup } from '../shared/groupType';
import { getManageability, IManageability } from '../shared/manageability';
import { friendlyError, isTimeout, requestResultText } from '../shared/errors';
import { diag } from '../shared/log';

interface IAlert {
  type: MessageBarType;
  text: string;
}

interface ICardState {
  membersLoading?: boolean;
  memberError?: string;
  members?: IUser[];
  ownersLoading?: boolean;
  ownersError?: string;
  owners?: IUser[];
  directoryQuery?: string;
  directoryLoading?: boolean;
  directoryResults?: IUser[];
  selectedUser?: IUser;
  processing?: boolean;
  processingMessage?: string;
  alert?: IAlert;
  confirmRemove?: IUser;
}

const ACTION_LABEL: { [key: string]: string } = {
  'Add Member': 'Add member',
  'Remove Member': 'Remove'
};

function initials(name: string): string {
  return (
    (name || '')
      .split(' ')
      .filter(Boolean)
      .slice(0, 2)
      .map((w: string) => w.charAt(0).toUpperCase())
      .join('') || '?'
  );
}

function applyOfficeScope(all: IOfficeGroup[], spec: string): IOfficeGroup[] {
  const wanted: string[] = (spec || '')
    .split(',')
    .map((s: string) => s.trim().toLowerCase())
    .filter(Boolean);
  if (wanted.length === 0) {
    return all;
  }
  return all.filter(
    (g: IOfficeGroup) =>
      wanted.indexOf((g.groupId || '').toLowerCase()) !== -1 ||
      wanted.indexOf((g.title || '').toLowerCase()) !== -1
  );
}

const AccountManagement: React.FunctionComponent<IAccountManagementProps> = (props: IAccountManagementProps) => {
  const spService: React.MutableRefObject<AccountManagementService> = React.useRef<AccountManagementService>(
    new AccountManagementService(props.context, props.listConfig)
  );
  const graphService: React.MutableRefObject<GraphService> = React.useRef<GraphService>(
    new GraphService(props.context)
  );
  const searchTimers: React.MutableRefObject<{ [id: number]: number }> = React.useRef<{ [id: number]: number }>({});

  const [groups, setGroups] = React.useState<IOfficeGroup[]>([]);
  const [authorizedCount, setAuthorizedCount] = React.useState<number>(0);
  const [selectedGroupId, setSelectedGroupId] = React.useState<number | undefined>(undefined);
  const [groupSearch, setGroupSearch] = React.useState<string>('');
  const [loading, setLoading] = React.useState<boolean>(true);
  const [topError, setTopError] = React.useState<string | undefined>(undefined);
  const [cards, setCards] = React.useState<{ [id: number]: ICardState }>({});
  const [recent, setRecent] = React.useState<IRequestSummary[]>([]);

  const currentUserKey: string = (props.context.pageContext.user.email || '').toLowerCase();

  const updateCard = (id: number, patch: ICardState): void => {
    setCards((prev: { [id: number]: ICardState }) => {
      const next: { [id: number]: ICardState } = { ...prev };
      next[id] = { ...(prev[id] || {}), ...patch };
      return next;
    });
  };

  React.useEffect(() => {
    diag('365 Account Management diagnostic React component mounted', {
      buildVersion: props.buildVersion,
      userDisplayName: props.userDisplayName,
      siteUrl: props.context.pageContext.web.absoluteUrl
    });
  }, []);

  const loadRecent = async (): Promise<void> => {
    try {
      setRecent(await spService.current.getRecentRequests(25));
    } catch (e) {
      /* recent history is best-effort */
    }
  };

  // Load authorized groups on mount.
  React.useEffect(() => {
    let cancelled: boolean = false;
    (async (): Promise<void> => {
      try {
        diag('365 Account Management diagnostic loading authorized groups');
        const loaded: IOfficeGroup[] = await spService.current.getAuthorizedGroups();
        if (!cancelled) {
          const scoped: IOfficeGroup[] = applyOfficeScope(loaded, props.visibleOffices);
          diag('365 Account Management diagnostic authorized groups loaded', {
            count: loaded.length,
            shownAfterScope: scoped.length
          });
          setAuthorizedCount(loaded.length);
          setGroups(scoped);
          setSelectedGroupId(scoped[0] ? scoped[0].id : undefined);
          loadRecent().catch(() => undefined);
        }
      } catch (err) {
        console.error('365 Account Management failed to load authorized groups.', { error: err });
        if (!cancelled) {
          setTopError(friendlyError(err));
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    })().catch((err: unknown) => {
      setTopError(friendlyError(err));
    });
    return () => {
      cancelled = true;
    };
  }, []);

  const loadMembers = async (group: IOfficeGroup): Promise<void> => {
    updateCard(group.id, { membersLoading: true, memberError: undefined, members: undefined });
    try {
      const members: IUser[] = await graphService.current.getGroupMembers(group.groupId);
      updateCard(group.id, { members: members, membersLoading: false });
    } catch (err) {
      console.error('365 Account Management failed to load group members.', { error: err });
      updateCard(group.id, { memberError: friendlyError(err), membersLoading: false });
    }
  };

  const loadOwners = async (group: IOfficeGroup): Promise<void> => {
    if (isSharePointGroup(group.groupId)) {
      return; // owners are an O365/Graph concept
    }
    updateCard(group.id, { ownersLoading: true, ownersError: undefined });
    try {
      const owners: IUser[] = await graphService.current.getGroupOwners(group.groupId);
      updateCard(group.id, { owners: owners, ownersLoading: false });
    } catch (err) {
      updateCard(group.id, { ownersError: friendlyError(err), ownersLoading: false });
    }
  };

  // Load members + owners for the expanded, manageable group.
  React.useEffect(() => {
    if (selectedGroupId === undefined) {
      return;
    }
    const group: IOfficeGroup | undefined = groups.filter((g: IOfficeGroup) => g.id === selectedGroupId)[0];
    if (!group || !getManageability(group).manageable) {
      return;
    }
    const card: ICardState = cards[selectedGroupId] || {};
    if (!card.membersLoading && !card.members) {
      loadMembers(group).catch(() => undefined);
    }
    if (!card.ownersLoading && !card.owners && !card.ownersError) {
      loadOwners(group).catch(() => undefined);
    }
  }, [selectedGroupId]);

  const runDirectorySearch = async (group: IOfficeGroup, term: string): Promise<void> => {
    try {
      const results: IUser[] = await graphService.current.searchUsers(term);
      // Exclude current members across both id-spaces by id AND upn/mail.
      const members: IUser[] = (cards[group.id] && cards[group.id].members) || [];
      const seen: Set<string> = new Set<string>();
      members.forEach((m: IUser) => {
        if (m.id) { seen.add(m.id.toLowerCase()); }
        if (m.userPrincipalName) { seen.add(m.userPrincipalName.toLowerCase()); }
        if (m.mail) { seen.add(m.mail.toLowerCase()); }
      });
      const filtered: IUser[] = results.filter(
        (u: IUser) =>
          !seen.has((u.id || '').toLowerCase()) &&
          !seen.has((u.userPrincipalName || '').toLowerCase()) &&
          !seen.has((u.mail || '').toLowerCase())
      );
      updateCard(group.id, { directoryResults: filtered, directoryLoading: false });
    } catch (err) {
      console.error('365 Account Management failed to search users.', { error: err });
      updateCard(group.id, {
        directoryResults: [],
        directoryLoading: false,
        alert: { type: MessageBarType.error, text: friendlyError(err) }
      });
    }
  };

  const onDirectoryChange = (group: IOfficeGroup, value: string): void => {
    updateCard(group.id, { directoryQuery: value, selectedUser: undefined });
    const term: string = (value || '').trim();
    if (searchTimers.current[group.id] !== undefined) {
      window.clearTimeout(searchTimers.current[group.id]);
    }
    if (term.length < 2) {
      updateCard(group.id, { directoryResults: [], directoryLoading: false });
      return;
    }
    updateCard(group.id, { directoryLoading: true });
    searchTimers.current[group.id] = window.setTimeout(() => {
      runDirectorySearch(group, term).catch((err: unknown) =>
        updateCard(group.id, {
          directoryLoading: false,
          alert: { type: MessageBarType.error, text: friendlyError(err) }
        })
      );
    }, 350);
  };

  const submit = async (group: IOfficeGroup, action: MembershipAction, member: IUser): Promise<void> => {
    updateCard(group.id, {
      processing: true,
      processingMessage: `Submitting ${ACTION_LABEL[action].toLowerCase()} request...`,
      alert: undefined,
      confirmRemove: undefined
    });
    try {
      if (isSharePointGroup(group.groupId)) {
        updateCard(group.id, { processingMessage: 'Updating SharePoint group...' });
        await spService.current.changeSharePointGroupMembership({ action, spGroupId: group.groupId, member });
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          selectedUser: undefined,
          directoryQuery: '',
          directoryResults: [],
          alert: { type: MessageBarType.success, text: 'Membership updated.' }
        });
      } else {
        const created = await spService.current.createMembershipRequest({ action, group, member });
        updateCard(group.id, { processingMessage: 'Waiting for Power Automate to finish...' });
        const result = await spService.current.pollRequest(
          created.id,
          (r) => updateCard(group.id, { processingMessage: `Request status: ${r.status}` }),
          props.pollTimeoutMs
        );
        const ok: boolean = result.status === 'Completed';
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          selectedUser: undefined,
          directoryQuery: '',
          directoryResults: [],
          alert: {
            type: ok ? MessageBarType.success : MessageBarType.error,
            text: requestResultText(result.resultMessage, result.authorizationResult, ok)
          }
        });
        loadRecent().catch(() => undefined);
      }
    } catch (err) {
      if (isTimeout(err)) {
        // The flow may still finish — present as a neutral "still processing" warning, not a failure.
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          selectedUser: undefined,
          alert: {
            type: MessageBarType.warning,
            text: 'Your request was submitted and is still processing. Refresh the members in a minute to confirm.'
          }
        });
        loadRecent().catch(() => undefined);
      } else {
        console.error('365 Account Management failed to submit membership request.', { error: err });
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          alert: { type: MessageBarType.error, text: friendlyError(err) }
        });
      }
    } finally {
      // Always refresh the member list — even on timeout, where the change may have landed.
      await loadMembers(group);
    }
  };

  const renderPersona = (user: IUser, inner: JSX.Element): JSX.Element => {
    const upn: string = user.userPrincipalName || user.mail || '';
    if (!upn) {
      return inner;
    }
    // ServiceScope is duplicated under @pnp's nested sp-core-library; structurally identical at runtime.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return <LivePersona upn={upn} serviceScope={props.context.serviceScope as any} template={inner} />;
  };

  const versionBadge: JSX.Element = (
    <div className={styles.diagnosticBadge}>365 Account Management v{props.buildVersion}</div>
  );

  if (loading) {
    return (
      <section className={styles.accountManagement}>
        {versionBadge}
        <Spinner label="Loading group access..." size={SpinnerSize.medium} />
      </section>
    );
  }

  // Friendly empty states instead of rendering nothing.
  if (!topError && groups.length === 0) {
    diag('365 Account Management diagnostic no authorized groups; web part hidden');
    const scopedOut: boolean = authorizedCount > 0;
    return (
      <section className={styles.accountManagement}>
        {versionBadge}
        <MessageBar messageBarType={MessageBarType.info}>
          {scopedOut
            ? 'None of the offices you are authorized to manage are configured to show on this page. Check the web part’s “Offices to show” setting.'
            : 'You are not currently authorized to manage any offices here. If you expect access, contact your administrator.'}
        </MessageBar>
      </section>
    );
  }

  const filtered: IOfficeGroup[] = groups.filter((g: IOfficeGroup) => {
    const t: string = groupSearch.toLowerCase();
    return (
      !t ||
      g.title.toLowerCase().indexOf(t) !== -1 ||
      (g.mail || '').toLowerCase().indexOf(t) !== -1 ||
      (g.siteTitle || '').toLowerCase().indexOf(t) !== -1
    );
  });

  return (
    <section className={styles.accountManagement}>
      {versionBadge}
      <div className={styles.header}>
        <h2>365 Account Management</h2>
        <p>Add or remove Microsoft 365 group members for groups you are authorized to manage.</p>
      </div>

      {topError && <MessageBar messageBarType={MessageBarType.error}>{topError}</MessageBar>}

      {!topError && (
        <>
          <div className={styles.toolbar}>
            <SearchBox
              placeholder="Search groups"
              value={groupSearch}
              onChange={(_: unknown, v?: string) => setGroupSearch(v || '')}
              className={styles.search}
              ariaLabel="Search the offices you manage"
            />
          </div>

          <div className={styles.grid}>
            {filtered.map((group: IOfficeGroup) => {
              const card: ICardState = cards[group.id] || {};
              const expanded: boolean = selectedGroupId === group.id;
              const manage: IManageability = getManageability(group);
              const recentForGroup: IRequestSummary[] = recent.filter(
                (r: IRequestSummary) => (r.groupId || '').toLowerCase() === (group.groupId || '').toLowerCase()
              );
              return (
                <article className={styles.groupCard} key={group.id}>
                  <button
                    className={styles.groupHeader}
                    type="button"
                    onClick={() => setSelectedGroupId(expanded ? undefined : group.id)}
                    aria-expanded={expanded}
                  >
                    <span className={styles.groupIcon} aria-hidden="true">
                      <Icon iconName="Group" />
                    </span>
                    <span className={styles.groupTitleBlock}>
                      <span className={styles.groupTitle}>{group.title}</span>
                      <span className={styles.groupMeta}>{group.mail || group.siteTitle || group.groupId}</span>
                    </span>
                    <Icon iconName={expanded ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
                  </button>

                  {expanded && (
                    <div className={styles.groupBody}>
                      <div className={styles.metaRow}>
                        {group.visibility && <span>{group.visibility}</span>}
                        {group.isTeamsConnected && <span>Teams connected</span>}
                        {group.siteTitle && <span>{group.siteTitle}</span>}
                      </div>

                      {!manage.manageable && (
                        <MessageBar messageBarType={MessageBarType.warning}>{manage.reason}</MessageBar>
                      )}

                      {card.alert && (
                        <MessageBar
                          messageBarType={card.alert.type}
                          onDismiss={() => updateCard(group.id, { alert: undefined })}
                        >
                          {card.alert.text}
                        </MessageBar>
                      )}

                      {card.confirmRemove && (
                        <MessageBar
                          messageBarType={MessageBarType.warning}
                          actions={
                            <div>
                              <PrimaryButton
                                text="Remove"
                                onClick={() => submit(group, 'Remove Member', card.confirmRemove as IUser)}
                              />
                              <DefaultButton text="Cancel" onClick={() => updateCard(group.id, { confirmRemove: undefined })} />
                            </div>
                          }
                        >
                          Remove <strong>{card.confirmRemove.displayName}</strong> from {group.title}?
                          {(card.confirmRemove.userPrincipalName || card.confirmRemove.mail || '').toLowerCase() === currentUserKey
                            ? ' This is your own access.'
                            : ''}
                          {(card.members ? card.members.length : 0) === 1 ? ' This is the last member of the group.' : ''}
                        </MessageBar>
                      )}

                      {card.processing && (
                        <div className={styles.processing} aria-live="polite">
                          <Spinner size={SpinnerSize.small} />
                          <span>{card.processingMessage}</span>
                        </div>
                      )}

                      {manage.manageable && (
                        <div className={styles.addArea}>
                          <SearchBox
                            placeholder="Search directory to add a member"
                            value={card.directoryQuery || ''}
                            ariaLabel={`Search the directory to add a member to ${group.title}`}
                            onChange={(_: unknown, v?: string) => onDirectoryChange(group, v || '')}
                            disabled={card.processing}
                          />
                          {card.directoryLoading && <Spinner size={SpinnerSize.small} label="Searching..." />}
                          {!!(card.directoryResults && card.directoryResults.length) && (
                            <div className={styles.directoryResults} role="listbox" aria-label="Directory results">
                              {card.directoryResults.map((u: IUser) => (
                                <button
                                  key={u.id}
                                  type="button"
                                  role="option"
                                  aria-selected={!!(card.selectedUser && card.selectedUser.id === u.id)}
                                  className={`${styles.personRow} ${
                                    card.selectedUser && card.selectedUser.id === u.id ? styles.personRowSelected : ''
                                  }`}
                                  onClick={() => updateCard(group.id, { selectedUser: u })}
                                >
                                  <span className={styles.avatar}>{initials(u.displayName)}</span>
                                  <span className={styles.memberDetails}>
                                    <strong>{u.displayName}</strong>
                                    <span>{u.jobTitle || u.mail || u.userPrincipalName}</span>
                                  </span>
                                </button>
                              ))}
                            </div>
                          )}
                          {card.selectedUser && (
                            <div className={styles.selectedUser}>
                              <span>
                                Add <strong>{card.selectedUser.displayName}</strong>
                              </span>
                              <PrimaryButton
                                text="Submit"
                                iconProps={{ iconName: 'AddFriend' }}
                                disabled={card.processing}
                                onClick={() => submit(group, 'Add Member', card.selectedUser as IUser)}
                              />
                            </div>
                          )}
                        </div>
                      )}

                      {!isSharePointGroup(group.groupId) && manage.manageable && (card.owners || card.ownersLoading || card.ownersError) && (
                        <div className={styles.subSection}>
                          <h3>Owners</h3>
                          {card.ownersLoading && <Spinner size={SpinnerSize.small} label="Loading owners..." />}
                          {card.ownersError && <span className={styles.emptyText}>Owners couldn’t be loaded.</span>}
                          {!card.ownersLoading && !card.ownersError && (
                            <div className={styles.memberList}>
                              {(card.owners || []).map((o: IUser) => (
                                <div className={styles.memberRow} key={`own-${o.id}`}>
                                  <div className={styles.personaWrap}>
                                    {renderPersona(
                                      o,
                                      <span className={styles.persona}>
                                        <span className={styles.avatar}>{initials(o.displayName)}</span>
                                        <span className={styles.memberDetails}>
                                          <strong>{o.displayName}</strong>
                                          <span>{o.jobTitle || o.mail || o.userPrincipalName}</span>
                                        </span>
                                      </span>
                                    )}
                                  </div>
                                </div>
                              ))}
                              {(card.owners ? card.owners.length : 0) === 0 && (
                                <p className={styles.emptyText}>No owners returned.</p>
                              )}
                            </div>
                          )}
                        </div>
                      )}

                      <div className={styles.membersHeader}>
                        <h3>Current members</h3>
                        <IconButton
                          iconProps={{ iconName: 'Refresh' }}
                          title="Refresh members"
                          ariaLabel="Refresh members"
                          disabled={card.processing || card.membersLoading}
                          onClick={() => {
                            loadMembers(group).catch(() => undefined);
                          }}
                        />
                      </div>

                      {card.memberError && <MessageBar messageBarType={MessageBarType.error}>{card.memberError}</MessageBar>}
                      {card.membersLoading && <Spinner label="Loading members..." size={SpinnerSize.small} />}

                      {!card.membersLoading && !card.memberError && (
                        <div className={styles.memberList}>
                          {(card.members || []).map((m: IUser) => (
                            <div className={styles.memberRow} key={m.id}>
                              <div className={styles.personaWrap}>
                                {renderPersona(
                                  m,
                                  <span className={styles.persona}>
                                    <span className={styles.avatar}>{initials(m.displayName)}</span>
                                    <span className={styles.memberDetails}>
                                      <strong>{m.displayName}</strong>
                                      <span>{m.jobTitle || m.mail || m.userPrincipalName}</span>
                                    </span>
                                  </span>
                                )}
                              </div>
                              {manage.manageable && (
                                <DefaultButton
                                  text="Remove"
                                  iconProps={{ iconName: 'RemoveFromShoppingList' }}
                                  disabled={card.processing}
                                  onClick={() => updateCard(group.id, { confirmRemove: m })}
                                />
                              )}
                            </div>
                          ))}
                          {(card.members ? card.members.length : 0) === 0 && (
                            <p className={styles.emptyText}>No members were returned for this group.</p>
                          )}
                        </div>
                      )}

                      {recentForGroup.length > 0 && (
                        <div className={styles.subSection}>
                          <h3>Your recent requests</h3>
                          <div className={styles.recentList}>
                            {recentForGroup.map((r: IRequestSummary) => (
                              <div className={styles.recentRow} key={`req-${r.id}`}>
                                <span
                                  className={`${styles.statusPill} ${
                                    r.status === 'Completed'
                                      ? styles.statusCompleted
                                      : r.status === 'Failed'
                                      ? styles.statusFailed
                                      : styles.statusPending
                                  }`}
                                >
                                  {r.status}
                                </span>
                                <span className={styles.memberDetails}>
                                  <strong>{r.action}: {r.memberDisplayName}</strong>
                                  <span>{r.resultMessage || r.requestedOn || ''}</span>
                                </span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </article>
              );
            })}
          </div>
        </>
      )}
    </section>
  );
};

export default AccountManagement;
