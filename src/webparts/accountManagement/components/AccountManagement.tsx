import * as React from 'react';
import {
  FluentProvider,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  SearchBox,
  Field,
  Textarea,
  Link,
  Button,
  Switch
} from '@fluentui/react-components';
import {
  ChevronUp20Regular,
  ChevronDown20Regular,
  PersonAdd20Regular,
  People20Regular,
  ArrowSync20Regular,
  Delete20Regular,
  Dismiss20Regular,
  Print20Regular,
  Warning20Regular,
  Clock16Regular,
  CheckmarkCircle16Regular,
  ErrorCircle16Regular
} from '@fluentui/react-icons';
import { buildFluentTheme } from './theme';
import styles from './AccountManagement.module.scss';
import { IAccountManagementProps } from './IAccountManagementProps';
import ListHealthPanel from './ListHealthPanel';
import { IListHealth } from '../models/listSchema';
import { ConfirmDialog } from './ConfirmDialog';
import { GraphService } from '../services/GraphService';
import { AccountManagementService } from '../services/AccountManagementService';
import { IOfficeGroup, IUser, IRequestSummary, ISitePermission, MembershipAction } from '../models/types';
import { isSharePointGroup, groupKindLabel } from '../shared/groupType';
import { getManageability, IManageability } from '../shared/manageability';
import { friendlyError, isTimeout, requestResultText } from '../shared/errors';
import { diag } from '../shared/log';

/** Fluent v9 MessageBar intent (replaces Fluent v8 MessageBarType). */
type AlertIntent = 'success' | 'error' | 'warning' | 'info';

interface IAlert {
  type: AlertIntent;
  text: string;
  /** The O365 request this banner is tracking, so it clears only when THAT request finishes. */
  requestId?: number;
}

interface ICardState {
  membersLoading?: boolean;
  memberError?: string;
  members?: IUser[];
  memberFilter?: string;
  justification?: string;
  ownersLoading?: boolean;
  ownersError?: string;
  owners?: IUser[];
  directoryQuery?: string;
  addMode?: 'people' | 'groups';
  directoryLoading?: boolean;
  directoryResults?: IUser[];
  directoryCapped?: boolean;
  selectedUser?: IUser;
  processing?: boolean;
  processingMessage?: string;
  alert?: IAlert;
  confirmRemove?: IUser;
  confirmRemoveAction?: MembershipAction;
  recentOpen?: boolean;
  activeTab?: string;
  sitePerms?: ISitePermission[];
  sitePermsLoading?: boolean;
}

const ACTION_LABEL: { [key: string]: string } = {
  'Add Member': 'Add member',
  'Remove Member': 'Remove',
  'Add Owner': 'Add owner',
  'Remove Owner': 'Remove owner'
};

// Owner actions target a group's owners collection (O365 only); member actions target members.
const isOwnerAction = (a: MembershipAction): boolean => a === 'Add Owner' || a === 'Remove Owner';
const isRemoveAction = (a: MembershipAction): boolean => a === 'Remove Member' || a === 'Remove Owner';
const inverseAction = (a: MembershipAction): MembershipAction =>
  a === 'Add Member' ? 'Remove Member' : a === 'Remove Member' ? 'Add Member' : a === 'Add Owner' ? 'Remove Owner' : 'Add Owner';

// Directory-search tuning (large-tenant friendly).
const MIN_SEARCH_CHARS: number = 3;
const SEARCH_DEBOUNCE_MS: number = 400;
const SEARCH_RESULT_CAP: number = 25;
// Show the member filter box only once a group is big enough to warrant it.
const MEMBER_FILTER_THRESHOLD: number = 5;
// Cap how many member rows render at once (large groups); the filter box finds anyone beyond this.
const MEMBER_RENDER_CAP: number = 200;
// A request still Pending after this long is shown as Error (the workflow failed or stalled).
const STALE_PENDING_MS: number = 5 * 60 * 1000;

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

// Deterministic tile color for a group (mimics the Entra/Teams default colored group tile).
function colorFor(name: string): string {
  const palette: string[] = ['#2564cf', '#0b6a0b', '#a4262c', '#8764b8', '#038387', '#ca5010', '#5c2e91', '#486991'];
  const n: string = name || '';
  let sum: number = 0;
  for (let i: number = 0; i < n.length; i++) {
    sum += n.charCodeAt(i);
  }
  return palette[sum % palette.length];
}

function formatDateTime(iso: string | undefined): string {
  if (!iso) {
    return '';
  }
  const d: Date = new Date(iso);
  return isNaN(d.getTime()) ? '' : d.toLocaleString();
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
  const [health, setHealth] = React.useState<IListHealth[] | undefined>(undefined);
  const [cards, setCards] = React.useState<{ [id: number]: ICardState }>({});
  const [recent, setRecent] = React.useState<IRequestSummary[]>([]);
  const [printing, setPrinting] = React.useState<boolean>(false);
  const [printNote, setPrintNote] = React.useState<string | undefined>(undefined);
  const [groupPhotos, setGroupPhotos] = React.useState<{ [id: number]: string }>({});

  // Fluent v9 theme (maps the SharePoint section theme onto v9 brand tokens).
  const theme = React.useMemo(() => buildFluentTheme(props.sectionTheme), [props.sectionTheme]);

  const currentUserKey: string = (props.context.pageContext.user.email || '').toLowerCase();

  const updateCard = (id: number, patch: ICardState): void => {
    setCards((prev: { [id: number]: ICardState }) => {
      const next: { [id: number]: ICardState } = { ...prev };
      next[id] = { ...(prev[id] || {}), ...patch };
      return next;
    });
  };

  // Reflect a just-confirmed add/remove in the local member list immediately. The Graph
  // /groups/{id}/members endpoint is eventually consistent, so re-reading it right after the
  // flow's write can still return the pre-change list for several seconds. Trust the confirmed
  // result here; the Refresh button re-queries Graph once it has caught up.
  const applyMemberChange = (id: number, action: MembershipAction, member: IUser): void => {
    // UPN/mail first so the key matches across id-spaces (Graph GUID vs SharePoint site-user id).
    const key = (u: IUser): string => (u.userPrincipalName || u.mail || u.id || '').toLowerCase();
    const owner: boolean = isOwnerAction(action);
    setCards((prev: { [id: number]: ICardState }) => {
      const card: ICardState = prev[id] || {};
      const existing: IUser[] = (owner ? card.owners : card.members) || [];
      let updated: IUser[];
      if (isRemoveAction(action)) {
        updated = existing.filter((m: IUser) => key(m) !== key(member));
      } else if (existing.some((m: IUser) => key(m) === key(member))) {
        updated = existing; // already present (idempotent add) — leave as-is
      } else {
        updated = existing.concat([member]);
      }
      const next: { [id: number]: ICardState } = { ...prev };
      next[id] = owner ? { ...card, owners: updated } : { ...card, members: updated };
      return next;
    });
  };

  // Clear a card's banner only if it still belongs to the given request, so a completed request
  // never clobbers a newer one's banner.
  const clearAlertFor = (id: number, requestId: number): void => {
    setCards((prev: { [id: number]: ICardState }) => {
      const card: ICardState = prev[id] || {};
      if (card.alert && card.alert.requestId === requestId) {
        const next: { [id: number]: ICardState } = { ...prev };
        next[id] = { ...card, alert: undefined };
        return next;
      }
      return prev;
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
      const rows: IRequestSummary[] = await spService.current.getRecentRequests(25);
      setRecent(rows);
      // Drop any "submitted / tracking" banner whose request has reached a terminal state — even if the
      // background poll already gave up on a slow flow — so the banner never lingers after completion.
      const terminal: Set<number> = new Set<number>(
        rows
          .filter((r: IRequestSummary) => r.status === 'Completed' || r.status === 'Failed')
          .map((r: IRequestSummary) => r.id)
      );
      setCards((prev: { [id: number]: ICardState }) => {
        let changed: boolean = false;
        const next: { [id: number]: ICardState } = { ...prev };
        Object.keys(prev).forEach((k: string) => {
          const card: ICardState = prev[Number(k)];
          if (
            card.alert &&
            card.alert.type === 'info' &&
            card.alert.requestId !== undefined &&
            terminal.has(card.alert.requestId)
          ) {
            next[Number(k)] = { ...card, alert: undefined };
            changed = true;
          }
        });
        return changed ? next : prev;
      });
    } catch {
      /* recent history is best-effort */
    }
  };

  // (Re)load authorized groups on mount and whenever the list config or office scope changes, so
  // property-pane edits take effect without a full page reload.
  const listConfigKey: string = JSON.stringify(props.listConfig);
  React.useEffect(() => {
    let cancelled: boolean = false;
    // Rebuild the service so it queries the (possibly renamed) lists, and reset per-card state.
    spService.current = new AccountManagementService(props.context, props.listConfig);
    setLoading(true);
    setTopError(undefined);
    setCards({});
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
          setSelectedGroupId(props.startCollapsed ? undefined : scoped[0] ? scoped[0].id : undefined);
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
  }, [listConfigKey, props.visibleOffices]);

  const loadMembers = async (group: IOfficeGroup): Promise<void> => {
    updateCard(group.id, { membersLoading: true, memberError: undefined, members: undefined, memberFilter: undefined });
    try {
      const members: IUser[] = await graphService.current.getGroupMembers(group.groupId, group.siteUrl);
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

  const loadSitePerms = async (group: IOfficeGroup): Promise<void> => {
    updateCard(group.id, { sitePermsLoading: true });
    try {
      const rows: ISitePermission[] = await spService.current.getGroupSitePermissions(group.groupId);
      updateCard(group.id, { sitePerms: rows, sitePermsLoading: false });
    } catch {
      updateCard(group.id, { sitePerms: [], sitePermsLoading: false });
    }
  };

  // Self-diagnosing schema check: probe the configured lists once so the UI can degrade
  // gracefully and (in edit mode) say exactly what to fix. Never blocks the main load.
  React.useEffect(() => {
    let cancelled: boolean = false;
    spService.current
      .checkListHealth()
      .then((h: IListHealth[]) => {
        if (!cancelled) {
          setHealth(h);
        }
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [
    props.listConfig.groupListTitle,
    props.listConfig.requestListTitle,
    props.listConfig.authorizedAdminsListTitle,
    props.listConfig.sitePermissionsListTitle,
    props.listConfig.listSiteUrl
  ]);

  // Load site permissions for any expanded group; members + owners for manageable ones.
  React.useEffect(() => {
    if (selectedGroupId === undefined) {
      return;
    }
    const group: IOfficeGroup | undefined = groups.filter((g: IOfficeGroup) => g.id === selectedGroupId)[0];
    if (!group) {
      return;
    }
    const card: ICardState = cards[selectedGroupId] || {};
    // Fetch the M365 group photo lazily, on expand (was eager for EVERY group — a throttling risk at 30+ groups).
    if (props.showGroupPhotos && !isSharePointGroup(group.groupId) && !groupPhotos[group.id]) {
      const gid: number = group.id;
      graphService.current
        .getGroupPhotoUrl(group.groupId)
        .then((url: string | undefined) => {
          if (url) {
            setGroupPhotos((prev: { [id: number]: string }) => ({ ...prev, [gid]: url }));
          }
        })
        .catch(() => undefined);
    }
    if (!card.sitePermsLoading && !card.sitePerms) {
      loadSitePerms(group).catch(() => undefined);
    }
    if (!getManageability(group).manageable) {
      return;
    }
    if (!card.membersLoading && !card.members) {
      loadMembers(group).catch(() => undefined);
    }
    if (!card.ownersLoading && !card.owners && !card.ownersError) {
      loadOwners(group).catch(() => undefined);
    }
  }, [selectedGroupId]);

  const runDirectorySearch = async (group: IOfficeGroup, term: string): Promise<void> => {
    try {
      const mode: string = (cards[group.id] && cards[group.id].addMode) || 'people';
      const results: IUser[] =
        mode === 'groups'
          ? await graphService.current.searchGroups(term)
          : await graphService.current.searchUsers(term);
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
      updateCard(group.id, {
        directoryResults: filtered,
        directoryCapped: results.length >= SEARCH_RESULT_CAP,
        directoryLoading: false
      });
    } catch (err) {
      console.error('365 Account Management failed to search users.', { error: err });
      updateCard(group.id, {
        directoryResults: [],
        directoryLoading: false,
        alert: { type: 'error', text: friendlyError(err) }
      });
    }
  };

  const onDirectoryChange = (group: IOfficeGroup, value: string): void => {
    updateCard(group.id, { directoryQuery: value, selectedUser: undefined });
    const term: string = (value || '').trim();
    if (searchTimers.current[group.id] !== undefined) {
      window.clearTimeout(searchTimers.current[group.id]);
    }
    if (term.length < MIN_SEARCH_CHARS) {
      updateCard(group.id, { directoryResults: [], directoryCapped: false, directoryLoading: false });
      return;
    }
    updateCard(group.id, { directoryLoading: true });
    searchTimers.current[group.id] = window.setTimeout(() => {
      runDirectorySearch(group, term).catch((err: unknown) =>
        updateCard(group.id, {
          directoryLoading: false,
          alert: { type: 'error', text: friendlyError(err) }
        })
      );
    }, SEARCH_DEBOUNCE_MS);
  };

  // Background tracker for a queued O365 request: polls the request item without blocking the card,
  // then reconciles — refreshes recent, and reverts the optimistic member change if the flow Failed.
  const trackRequest = (
    group: IOfficeGroup,
    action: MembershipAction,
    member: IUser,
    requestId: number
  ): void => {
    spService.current
      .pollRequest(requestId, () => undefined, props.pollTimeoutMs)
      .then((result) => {
        if (result.status === 'Failed') {
          // Undo the optimistic change, name the member so a queued failure is unambiguous, and
          // reconcile the list against the server.
          applyMemberChange(group.id, inverseAction(action), member);
          updateCard(group.id, {
            alert: {
              type: 'error',
              text: `${member.displayName}: ${requestResultText(result.resultMessage, result.authorizationResult, false)}`,
              requestId
            }
          });
          (isOwnerAction(action) ? loadOwners(group) : loadMembers(group)).catch(() => undefined);
        } else {
          // Completed — drop this request's "submitted / tracking" banner if it's still showing.
          clearAlertFor(group.id, requestId);
        }
        loadRecent().catch(() => undefined);
      })
      .catch(() => {
        // Timed out or errored mid-poll — re-read the affected list so we stop trusting the optimistic state,
        // and refresh recent so the (now possibly stale-Pending -> Error) status shows.
        (isOwnerAction(action) ? loadOwners(group) : loadMembers(group)).catch(() => undefined);
        loadRecent().catch(() => undefined);
      });
  };

  const submit = async (
    group: IOfficeGroup,
    action: MembershipAction,
    member: IUser,
    justification?: string
  ): Promise<void> => {
    updateCard(group.id, {
      processing: true,
      processingMessage: `Submitting ${ACTION_LABEL[action].toLowerCase()} request...`,
      alert: undefined,
      confirmRemove: undefined
    });
    try {
      if (isSharePointGroup(group.groupId)) {
        updateCard(group.id, { processingMessage: 'Updating SharePoint group...' });
        await spService.current.changeSharePointGroupMembership({ action, spGroupId: group.groupId, member, siteUrl: group.siteUrl });
        // Write a who/why audit record for the direct SP change; surface it if the audit write fails.
        const audited: boolean = await spService.current.recordCompletedChange({ action, group, member, justification });
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          selectedUser: undefined,
          directoryQuery: '',
          directoryResults: [],
          justification: undefined,
          alert: audited
            ? { type: 'success', text: 'Membership updated.' }
            : {
                type: 'warning',
                text: 'Membership updated, but the audit record could not be written — notify an administrator.'
              }
        });
        loadRecent().catch(() => undefined);
        // SharePoint REST is read-after-write consistent, so re-reading reflects the change.
        await loadMembers(group);
      } else {
        // Queue model: file the Pending request, reflect the change optimistically, then re-enable the
        // card right away and track the flow's result in the background so more changes can be queued.
        const created = await spService.current.createMembershipRequest({ action, group, member, justification });
        applyMemberChange(group.id, action, member);
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          selectedUser: undefined,
          directoryQuery: '',
          directoryResults: [],
          justification: undefined,
          recentOpen: true,
          alert: {
            type: 'info',
            text: `Request submitted for ${member.displayName}. Tracking it in the Recent Requests tab.`,
            requestId: created.id
          }
        });
        loadRecent().catch(() => undefined);
        trackRequest(group, action, member, created.id);
      }
    } catch (err) {
      if (isTimeout(err)) {
        // The flow may still finish — present as a neutral "still processing" warning, not a failure.
        // Don't re-read here: the change may or may not have landed and Graph is lagging either way;
        // the message tells the user to use Refresh once it settles.
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          selectedUser: undefined,
          alert: {
            type: 'warning',
            text: 'Your request was submitted and is still processing. Refresh the members in a minute to confirm.'
          }
        });
        loadRecent().catch(() => undefined);
      } else {
        console.error('365 Account Management failed to submit membership request.', { error: err });
        updateCard(group.id, {
          processing: false,
          processingMessage: undefined,
          alert: { type: 'error', text: friendlyError(err) }
        });
        // A Pending request row may have been written before polling failed — keep history in sync.
        loadRecent().catch(() => undefined);
        // The change did not go through, so the current list is still accurate; re-read to be safe.
        await loadMembers(group);
      }
    }
  };

  interface IPrintSection {
    group: IOfficeGroup;
    members: IUser[];
    owners: IUser[];
  }

  // Print View: gather members (and O365 owners) for every manageable group, then open a clean,
  // self-contained print document in a new window (the web part can't restyle the whole SP page).
  const printAll = async (): Promise<void> => {
    setPrintNote(undefined);
    // Open the print window synchronously inside the click gesture. Opening it AFTER the awaited Graph
    // fetches would run outside the user gesture and get blocked as an unsolicited pop-up.
    const win: Window | null = window.open('', '_blank');
    if (!win) {
      setPrintNote('Your browser blocked the print window. Allow pop-ups for this site and try again.');
      return;
    }
    win.document.write(
      '<!doctype html><title>Preparing…</title><body style="font:14px Segoe UI,Arial,sans-serif;color:#444;padding:24px">Preparing the membership report…</body>'
    );
    setPrinting(true);
    try {
      const sections: IPrintSection[] = [];
      for (const g of groups) {
        if (!getManageability(g).manageable) {
          continue;
        }
        let members: IUser[] | undefined = cards[g.id] && cards[g.id].members;
        if (!members) {
          try {
            members = await graphService.current.getGroupMembers(g.groupId, g.siteUrl);
          } catch {
            members = [];
          }
        }
        let owners: IUser[] = [];
        if (!isSharePointGroup(g.groupId)) {
          owners = (cards[g.id] && cards[g.id].owners) || [];
          if (owners.length === 0) {
            try {
              owners = await graphService.current.getGroupOwners(g.groupId);
            } catch {
              owners = [];
            }
          }
        }
        sections.push({ group: g, members: members || [], owners: owners });
      }
      writePrintDocument(win, sections);
    } catch {
      win.close();
      setPrintNote('Could not build the print view. Please try again.');
    } finally {
      setPrinting(false);
    }
  };

  const writePrintDocument = (win: Window, sections: IPrintSection[]): void => {
    const escapes: { [k: string]: string } = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
    const esc = (s: string | undefined): string => (s || '').replace(/[&<>"]/g, (c: string) => escapes[c]);
    const who: string = props.userDisplayName || props.context.pageContext.user.displayName || '';
    const stamp: string = new Date().toLocaleString();
    const body: string = sections
      .map((sec: IPrintSection) => {
        const rows: string =
          sec.members
            .map(
              (m: IUser) =>
                `<tr><td>${esc(m.displayName)}</td><td>${esc(m.mail || m.userPrincipalName)}</td><td>${esc(m.jobTitle) || '&mdash;'}</td></tr>`
            )
            .join('') || '<tr><td colspan="3">No members.</td></tr>';
        const owners: string = sec.owners.length
          ? `<p class="owners"><strong>Owners:</strong> ${sec.owners.map((o: IUser) => esc(o.displayName)).join(', ')}</p>`
          : '';
        return (
          `<section><h2>${esc(sec.group.title)}</h2>` +
          `<p class="meta">${esc(sec.group.mail || sec.group.siteTitle || groupKindLabel(sec.group.groupId))} &middot; ` +
          `${sec.members.length} member${sec.members.length === 1 ? '' : 's'}</p>${owners}` +
          `<table><thead><tr><th>Name</th><th>Email</th><th>Title</th></tr></thead><tbody>${rows}</tbody></table></section>`
        );
      })
      .join('');
    const html: string =
      '<!doctype html><html><head><meta charset="utf-8"><title>365 Account Management — Membership</title><style>' +
      'body{font-family:Segoe UI,Arial,sans-serif;color:#222;margin:24px;}' +
      'h1{font-size:20px;margin:0 0 4px;}h2{font-size:16px;margin:18px 0 4px;}' +
      'header{border-bottom:2px solid #ddd;padding-bottom:8px;margin-bottom:8px;}' +
      '.sub{color:#666;font-size:12px;margin:0;}.meta{color:#666;font-size:12px;margin:0 0 6px;}' +
      '.owners{font-size:12px;margin:0 0 6px;}' +
      'table{border-collapse:collapse;width:100%;font-size:12px;}' +
      'th,td{border:1px solid #ddd;padding:4px 8px;text-align:left;}th{background:#f3f3f3;}' +
      'section{page-break-inside:avoid;}@media print{body{margin:0;}}' +
      '</style></head><body>' +
      `<header><h1>Group Membership</h1><p class="sub">Printed by ${esc(who)} on ${esc(stamp)}</p></header>` +
      (body || '<p>No groups to print.</p>') +
      '<script>window.onload=function(){window.print();}</script></body></html>';
    win.document.open();
    win.document.write(html);
    win.document.close();
  };

  // LivePersona (hover profile card) was removed to drop the @pnp/spfx-controls-react dependency;
  // each row already shows the member's name, email, and title.
  const renderPersona = (inner: JSX.Element): JSX.Element => inner;

  const versionBadge: JSX.Element = (
    <div className={styles.diagnosticBadge}>365 Account Management v{props.buildVersion}</div>
  );

  // A missing required list is explained by the setup panel, so don't also show the raw
  // "a required SharePoint list was not found" error on top of it.
  const missingRequiredList: boolean =
    !!health && health.filter((h: IListHealth) => !h.exists && !h.optional && !h.error).length > 0;

  if (loading) {
    return (
      <FluentProvider theme={theme} style={{ background: 'transparent' }}>
        <section className={styles.accountManagement}>
          {versionBadge}
          <Spinner label="Loading group access..." size="medium" />
        </section>
      </FluentProvider>
    );
  }

  // Friendly empty states instead of rendering nothing.
  if (!topError && groups.length === 0) {
    diag('365 Account Management diagnostic no authorized groups; web part hidden');
    const scopedOut: boolean = authorizedCount > 0;
    return (
      <FluentProvider theme={theme} style={{ background: 'transparent' }}>
        <section className={styles.accountManagement}>
          {versionBadge}
          {health && <ListHealthPanel health={health} isEditMode={props.isEditMode} />}
          <MessageBar intent="info" layout="multiline">
            <MessageBarBody>
              {scopedOut
                ? 'None of the offices you are authorized to manage are configured to show on this page. Check the web part’s “Offices to show” setting.'
                : 'You are not currently authorized to manage any offices here. If you expect access, contact your administrator.'}
            </MessageBarBody>
          </MessageBar>
        </section>
      </FluentProvider>
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
    <FluentProvider theme={theme} style={{ background: 'transparent' }}>
      <section className={styles.accountManagement}>
        {versionBadge}
        <div className={styles.header}>
          <h2>{props.description || '365 Account Management'}</h2>
          <p>{props.introText || 'Add or remove Microsoft 365 group members for groups you are authorized to manage.'}</p>
          {props.helpText && (
            <p className={styles.helpLine}>
              {props.helpUrl ? (
                <Link href={props.helpUrl} target="_blank">
                  {props.helpText}
                </Link>
              ) : (
                props.helpText
              )}
            </p>
          )}
        </div>

        {health && <ListHealthPanel health={health} isEditMode={props.isEditMode} />}

        {topError && !missingRequiredList && (
          <MessageBar intent="error" politeness="assertive" layout="multiline">
            <MessageBarBody>{topError}</MessageBarBody>
          </MessageBar>
        )}

        {!topError && (
          <>
            <div className={styles.toolbar}>
              <SearchBox
                placeholder="Search groups"
                value={groupSearch}
                onChange={(_, data) => setGroupSearch(data.value || '')}
                className={styles.search}
                aria-label="Search the offices you manage"
              />
              <Button
                appearance="secondary"
                icon={<Print20Regular />}
                disabled={printing}
                onClick={() => {
                  printAll().catch(() => undefined);
                }}
              >
                {printing ? 'Preparing…' : 'Print all'}
              </Button>
            </div>
            {printNote && (
              <MessageBar intent="warning" layout="multiline">
                <MessageBarBody>{printNote}</MessageBarBody>
                <MessageBarActions
                  containerAction={
                    <Button
                      appearance="transparent"
                      aria-label="Dismiss"
                      icon={<Dismiss20Regular />}
                      onClick={() => setPrintNote(undefined)}
                    />
                  }
                />
              </MessageBar>
            )}

            <div className={styles.grid}>
              {filtered.map((group: IOfficeGroup) => {
                const card: ICardState = cards[group.id] || {};
                const expanded: boolean = selectedGroupId === group.id;
                const manage: IManageability = getManageability(group);
                const recentForGroup: IRequestSummary[] = recent.filter(
                  (r: IRequestSummary) => (r.groupId || '').toLowerCase() === (group.groupId || '').toLowerCase()
                );
                const memberFilterText: string = (card.memberFilter || '').trim().toLowerCase();
                const visibleMembers: IUser[] = (card.members || []).filter(
                  (m: IUser) =>
                    !memberFilterText ||
                    (m.displayName || '').toLowerCase().indexOf(memberFilterText) !== -1 ||
                    (m.mail || '').toLowerCase().indexOf(memberFilterText) !== -1 ||
                    (m.userPrincipalName || '').toLowerCase().indexOf(memberFilterText) !== -1 ||
                    (m.jobTitle || '').toLowerCase().indexOf(memberFilterText) !== -1
                );
                const justificationMissing: boolean =
                  props.requireJustification && !(card.justification || '').trim();
                return (
                  <article className={styles.groupCard} key={group.id}>
                    <button
                      className={styles.groupHeader}
                      type="button"
                      onClick={() => setSelectedGroupId(expanded ? undefined : group.id)}
                      aria-expanded={expanded}
                    >
                      <span className={styles.groupIcon} aria-hidden="true">
                        {groupPhotos[group.id] ? (
                          <img src={groupPhotos[group.id]} alt="" className={styles.groupPhoto} />
                        ) : (
                          <span className={styles.groupInitials} style={{ backgroundColor: colorFor(group.title) }}>
                            {initials(group.title)}
                          </span>
                        )}
                      </span>
                      <span className={styles.groupTitleBlock}>
                        <span className={styles.groupTitle}>{group.title}</span>
                        <span className={styles.groupMeta}>
                          {groupKindLabel(group.groupId)}
                          {group.mail || group.siteTitle ? ` · ${group.mail || group.siteTitle}` : ''}
                        </span>
                      </span>
                      {expanded ? (
                        <ChevronUp20Regular className={styles.chevron} />
                      ) : (
                        <ChevronDown20Regular className={styles.chevron} />
                      )}
                    </button>

                    {expanded && (
                      <div className={styles.groupBody}>
                        <div className={styles.metaRow}>
                          {group.visibility && <span>{group.visibility}</span>}
                          {group.isTeamsConnected && <span>Teams connected</span>}
                          {group.siteTitle && <span>{group.siteTitle}</span>}
                        </div>


                        {!manage.manageable && (
                          <MessageBar intent="warning" layout="multiline">
                            <MessageBarBody>{manage.reason}</MessageBarBody>
                          </MessageBar>
                        )}

                        {card.alert && (
                          <MessageBar intent={card.alert.type} politeness={card.alert.type === 'error' ? 'assertive' : 'polite'} layout="multiline">
                            <MessageBarBody>{card.alert.text}</MessageBarBody>
                            <MessageBarActions
                              containerAction={
                                <Button
                                  appearance="transparent"
                                  aria-label="Dismiss"
                                  icon={<Dismiss20Regular />}
                                  onClick={() => updateCard(group.id, { alert: undefined })}
                                />
                              }
                            />
                          </MessageBar>
                        )}

                        <ConfirmDialog
                          open={!!card.confirmRemove}
                          ariaLabel={
                            card.confirmRemove ? `Confirm removing ${card.confirmRemove.displayName}` : 'Confirm removal'
                          }
                          onDismiss={() => updateCard(group.id, { confirmRemove: undefined })}
                        >
                          {card.confirmRemove && (
                            <React.Fragment>
                              <div className={styles.confirmText}>
                                <Warning20Regular className={styles.confirmIcon} />
                                <span>
                                  Remove {card.confirmRemove.isGroup ? 'the group ' : ''}
                                  <strong>{card.confirmRemove.displayName}</strong>{' '}
                                  {card.confirmRemoveAction === 'Remove Owner' ? 'as an owner of' : 'from'} {group.title}?
                                  {(card.confirmRemove.userPrincipalName || card.confirmRemove.mail || '').toLowerCase() ===
                                  currentUserKey
                                    ? ' This is your own access.'
                                    : ''}
                                  {card.confirmRemoveAction === 'Remove Owner'
                                    ? (card.owners ? card.owners.length : 0) === 1
                                      ? ' This is the last owner of the group.'
                                      : ''
                                    : (card.members ? card.members.length : 0) === 1
                                    ? ' This is the last member of the group.'
                                    : ''}
                                </span>
                              </div>
                              {props.requireJustification && (
                                <Field label="Reason for this change" required>
                                  <Textarea
                                    rows={2}
                                    value={card.justification || ''}
                                    placeholder="Why are you making this change? (recorded with the request)"
                                    disabled={card.processing}
                                    onChange={(_, data) => updateCard(group.id, { justification: data.value || '' })}
                                  />
                                </Field>
                              )}
                              <div className={styles.confirmActions}>
                                <Button
                                  appearance="primary"
                                  className={justificationMissing ? undefined : styles.dangerButton}
                                  disabled={justificationMissing}
                                  onClick={() => submit(group, card.confirmRemoveAction || 'Remove Member', card.confirmRemove as IUser, card.justification)}
                                >
                                  Remove
                                </Button>
                                <Button onClick={() => updateCard(group.id, { confirmRemove: undefined })}>Cancel</Button>
                              </div>
                            </React.Fragment>
                          )}
                        </ConfirmDialog>

                        {card.processing && (
                          <div className={styles.processing} aria-live="polite">
                            <Spinner size="small" />
                            <span>{card.processingMessage}</span>
                          </div>
                        )}

                        {manage.manageable && (
                          <div className={styles.addArea}>
                            {isSharePointGroup(group.groupId) && (
                              <Switch
                                label="Add a group instead of a person"
                                checked={card.addMode === 'groups'}
                                disabled={card.processing}
                                onChange={(_, data) =>
                                  updateCard(group.id, {
                                    addMode: data.checked ? 'groups' : 'people',
                                    directoryQuery: '',
                                    directoryResults: [],
                                    selectedUser: undefined
                                  })
                                }
                              />
                            )}
                            <SearchBox
                              placeholder={
                                card.addMode === 'groups'
                                  ? 'Search for a group to add (type 3+ letters)'
                                  : 'Search directory to add a member (type 3+ letters)'
                              }
                              value={card.directoryQuery || ''}
                              aria-label={
                                card.addMode === 'groups'
                                  ? `Search for a group to add to ${group.title}`
                                  : `Search the directory to add a member to ${group.title}`
                              }
                              onChange={(_, data) => onDirectoryChange(group, data.value || '')}
                              disabled={card.processing}
                            />
                            {card.directoryLoading && <Spinner size="small" label="Searching..." />}
                            {card.directoryCapped && !card.directoryLoading && (
                              <p className={styles.emptyText}>Showing the first 25 matches — keep typing to narrow.</p>
                            )}
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
                                    <span className={styles.avatar}>
                                      {u.isGroup ? <People20Regular /> : initials(u.displayName)}
                                    </span>
                                    <span className={styles.memberDetails}>
                                      <strong>{u.displayName}</strong>
                                      <span>
                                        {u.isGroup
                                          ? u.mail
                                            ? `Group · ${u.mail}`
                                            : 'Group'
                                          : u.jobTitle || u.mail || u.userPrincipalName}
                                      </span>
                                    </span>
                                  </button>
                                ))}
                              </div>
                            )}
                            {card.selectedUser && (
                              <>
                                {props.requireJustification && (
                                  <Field label="Reason for this change" required>
                                    <Textarea
                                      rows={2}
                                      value={card.justification || ''}
                                      placeholder="Why are you making this change? (recorded with the request)"
                                      disabled={card.processing}
                                      onChange={(_, data) => updateCard(group.id, { justification: data.value || '' })}
                                    />
                                  </Field>
                                )}
                                <div className={styles.selectedUser}>
                                  <span>
                                    Add {card.selectedUser.isGroup ? 'the group ' : ''}
                                    <strong>{card.selectedUser.displayName}</strong>
                                  </span>
                                  <div className={styles.selectedUserActions}>
                                    <Button
                                      appearance="primary"
                                      icon={<PersonAdd20Regular />}
                                      disabled={card.processing || justificationMissing}
                                      onClick={() => submit(group, 'Add Member', card.selectedUser as IUser, card.justification)}
                                    >
                                      {isSharePointGroup(group.groupId) ? 'Submit' : 'Add as member'}
                                    </Button>
                                    {!isSharePointGroup(group.groupId) && !card.selectedUser.isGroup && (
                                      <Button
                                        disabled={card.processing || justificationMissing}
                                        onClick={() => submit(group, 'Add Owner', card.selectedUser as IUser, card.justification)}
                                      >
                                        Add as owner
                                      </Button>
                                    )}
                                  </div>
                                </div>
                              </>
                            )}
                          </div>
                        )}

                        {(() => {
                          const showOwners: boolean = !isSharePointGroup(group.groupId) && manage.manageable && !!(card.owners || card.ownersLoading || card.ownersError);
                          const showSites: boolean = !!(card.sitePerms && card.sitePerms.length);
                          const showRecent: boolean = recentForGroup.length > 0;
                          const keys: string[] = ['members', showOwners ? 'owners' : '', showSites ? 'sites' : '', showRecent ? 'recent' : ''].filter((k: string) => !!k);
                          const active: string = keys.indexOf(card.activeTab || '') !== -1 ? (card.activeTab as string) : 'members';
                          return (
                            <div className={styles.tabs}>
                              <div className={styles.tabStrip} role="tablist" aria-label="Group details">
                                <button type="button" role="tab" aria-selected={active === 'members'} className={`${styles.tab} ${active === 'members' ? styles.tabActive : ''}`} onClick={() => updateCard(group.id, { activeTab: 'members' })}>
                                  Members{card.members ? ` (${card.members.length})` : ''}
                                </button>
                                {showOwners && (
                                  <button type="button" role="tab" aria-selected={active === 'owners'} className={`${styles.tab} ${active === 'owners' ? styles.tabActive : ''}`} onClick={() => updateCard(group.id, { activeTab: 'owners' })}>
                                    Owners{card.owners ? ` (${card.owners.length})` : ''}
                                  </button>
                                )}
                                {showSites && (
                                  <button type="button" role="tab" aria-selected={active === 'sites'} className={`${styles.tab} ${active === 'sites' ? styles.tabActive : ''}`} onClick={() => updateCard(group.id, { activeTab: 'sites' })}>
                                    Used On{card.sitePerms ? ` (${card.sitePerms.length})` : ''}
                                  </button>
                                )}
                                {showRecent && (
                                  <button type="button" role="tab" aria-selected={active === 'recent'} className={`${styles.tab} ${active === 'recent' ? styles.tabActive : ''}`} onClick={() => updateCard(group.id, { activeTab: 'recent' })}>
                                    Recent Requests ({recentForGroup.length})
                                  </button>
                                )}
                              </div>
                              <div className={styles.tabPanel} role="tabpanel">
                                <div hidden={active !== 'members'}>
                        <div className={styles.membersHeader}>
                          <h3>Members{card.members ? ` (${card.members.length})` : ''}</h3>
                          <Button
                            appearance="transparent"
                            icon={<ArrowSync20Regular />}
                            title="Refresh members"
                            aria-label="Refresh members"
                            disabled={card.processing || card.membersLoading}
                            onClick={() => {
                              loadMembers(group).catch(() => undefined);
                            }}
                          />
                        </div>

                        {!card.membersLoading &&
                          !card.memberError &&
                          (card.members ? card.members.length : 0) > MEMBER_FILTER_THRESHOLD && (
                            <SearchBox
                              placeholder="Filter members"
                              value={card.memberFilter || ''}
                              aria-label={`Filter the members of ${group.title}`}
                              onChange={(_, data) => updateCard(group.id, { memberFilter: data.value || '' })}
                              className={styles.search}
                            />
                          )}

                        {card.memberError && (
                          <MessageBar intent="error" politeness="assertive" layout="multiline">
                            <MessageBarBody>{card.memberError}</MessageBarBody>
                          </MessageBar>
                        )}
                        {card.membersLoading && <Spinner label="Loading members..." size="small" />}

                        {!card.membersLoading && !card.memberError && (
                          <div className={styles.memberList}>
                            {visibleMembers.slice(0, MEMBER_RENDER_CAP).map((m: IUser) => (
                              <div className={styles.memberRow} key={m.id}>
                                <div className={styles.personaWrap}>
                                  {renderPersona(
                                    <span className={styles.persona}>
                                      <span className={styles.avatar}>
                                        {m.isGroup ? <People20Regular /> : initials(m.displayName)}
                                      </span>
                                      <span className={styles.memberDetails}>
                                        <strong>{m.displayName}</strong>
                                        <span title={m.isGroup ? (m.mail ? `Group · ${m.mail}` : 'Group') : (m.jobTitle || m.mail || m.userPrincipalName)}>
                                          {m.isGroup
                                            ? m.mail
                                              ? `Group · ${m.mail}`
                                              : 'Group'
                                            : m.jobTitle || m.mail || m.userPrincipalName}
                                        </span>
                                      </span>
                                    </span>
                                  )}
                                </div>
                                {manage.manageable && (
                                  <Button
                                    icon={<Delete20Regular />}
                                    disabled={card.processing}
                                    onClick={() => updateCard(group.id, { confirmRemove: m, confirmRemoveAction: 'Remove Member' })}
                                  >
                                    Remove
                                  </Button>
                                )}
                              </div>
                            ))}
                            {visibleMembers.length > MEMBER_RENDER_CAP && (
                              <p className={styles.emptyText}>
                                Showing the first {MEMBER_RENDER_CAP} of {visibleMembers.length}. Use the filter to find a specific person.
                              </p>
                            )}
                            {visibleMembers.length === 0 && (
                              <p className={styles.emptyText}>
                                {(card.members ? card.members.length : 0) === 0
                                  ? 'No members were returned for this group.'
                                  : 'No members match your filter.'}
                              </p>
                            )}
                          </div>
                        )}
                                </div>
                                <div hidden={active !== 'owners'}>
                        {!isSharePointGroup(group.groupId) && manage.manageable && (card.owners || card.ownersLoading || card.ownersError) && (
                          <div className={styles.subSection}>
                            <h3>Owners</h3>
                            {card.ownersLoading && <Spinner size="small" label="Loading owners..." />}
                            {card.ownersError && <span className={styles.emptyText}>Owners couldn’t be loaded.</span>}
                            {!card.ownersLoading && !card.ownersError && (
                              <div className={styles.memberList}>
                                {(card.owners || []).map((o: IUser) => (
                                  <div className={styles.memberRow} key={`own-${o.id}`}>
                                    <div className={styles.personaWrap}>
                                      {renderPersona(
                                        <span className={styles.persona}>
                                          <span className={styles.avatar}>{initials(o.displayName)}</span>
                                          <span className={styles.memberDetails}>
                                            <strong>{o.displayName}</strong>
                                            <span title={o.jobTitle || o.mail || o.userPrincipalName}>{o.jobTitle || o.mail || o.userPrincipalName}</span>
                                          </span>
                                        </span>
                                      )}
                                    </div>
                                    {manage.manageable && (
                                      <Button
                                        icon={<Delete20Regular />}
                                        disabled={card.processing}
                                        onClick={() => updateCard(group.id, { confirmRemove: o, confirmRemoveAction: 'Remove Owner' })}
                                      >
                                        Remove
                                      </Button>
                                    )}
                                  </div>
                                ))}
                                {(card.owners ? card.owners.length : 0) === 0 && (
                                  <p className={styles.emptyText}>No owners returned.</p>
                                )}
                              </div>
                            )}
                          </div>
                        )}
                                </div>
                                <div hidden={active !== 'sites'}>
                        {card.sitePerms && card.sitePerms.length > 0 && (
                          <div className={styles.subSection}>
                            <h3>Used on these sites</h3>
                            <div className={styles.sitePermList}>
                              {card.sitePerms.map((s: ISitePermission, i: number) => (
                                <div className={styles.sitePermRow} key={`sp-${i}`}>
                                  <span className={styles.sitePermSite}>
                                    {s.siteUrl ? (
                                      <Link href={s.siteUrl} target="_blank">
                                        {s.siteName}
                                      </Link>
                                    ) : (
                                      s.siteName
                                    )}
                                  </span>
                                  {s.permission && <span className={styles.sitePermLevel}>{s.permission}</span>}
                                </div>
                              ))}
                            </div>
                          </div>
                        )}
                                </div>
                                <div hidden={active !== 'recent'}>
                              <div className={styles.recentScroll}>
                                <div className={styles.recentList}>
                                  {recentForGroup.map((r: IRequestSummary) => {
                                    const edited: string = formatDateTime(r.modified);
                                    const submittedMs: number = r.requestedOn ? Date.parse(r.requestedOn) : NaN;
                                    const stalled: boolean =
                                      r.status === 'Pending' &&
                                      !isNaN(submittedMs) &&
                                      Date.now() - submittedMs > STALE_PENDING_MS;
                                    const displayStatus: string = stalled ? 'Error' : r.status;
                                    const detail: string = [
                                      r.resultMessage,
                                      stalled ? 'No response from the workflow — refresh to confirm' : '',
                                      edited ? `edited ${edited}` : ''
                                    ]
                                      .filter(Boolean)
                                      .join(' · ');
                                    return (
                                      <div className={styles.recentRow} key={`req-${r.id}`}>
                                        <span
                                          className={`${styles.statusPill} ${
                                            displayStatus === 'Completed'
                                              ? styles.statusCompleted
                                              : displayStatus === 'Failed' || displayStatus === 'Error'
                                              ? styles.statusFailed
                                              : styles.statusPending
                                          }`}
                                        >
                                          {displayStatus === 'Completed' ? (
                                            <CheckmarkCircle16Regular />
                                          ) : displayStatus === 'Failed' || displayStatus === 'Error' ? (
                                            <ErrorCircle16Regular />
                                          ) : (
                                            <Clock16Regular />
                                          )}
                                          {displayStatus}
                                        </span>
                                        <span className={styles.memberDetails}>
                                          <strong>{r.action}: {r.memberDisplayName}</strong>
                                          <span>{detail}</span>
                                        </span>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                                </div>
                              </div>
                            </div>
                          );
                        })()}
                      </div>
                    )}
                  </article>
                );
              })}
            </div>
          </>
        )}
      </section>
    </FluentProvider>
  );
};

export default AccountManagement;
