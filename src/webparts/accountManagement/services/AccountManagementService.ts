import { diag } from '../shared/log';
import { Guid } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IOfficeGroup,
  IMembershipRequest,
  IMembershipChangeInput,
  IRequestSummary,
  ICurrentUser,
  ISitePermission,
  IUser,
  MembershipAction
} from '../models/types';
import { toMessage } from '../shared/errors';

/* eslint-disable @typescript-eslint/no-explicit-any */

const DIAG: string = '365 Account Management diagnostic';

export interface IListConfig {
  requestListTitle: string;
  groupListTitle: string;
  authorizedAdminsListTitle: string;
  sitePermissionsListTitle: string;
  /** Optional: absolute URL of the site that holds the four lists. Blank = the current page's site. */
  listSiteUrl?: string;
}

export const DEFAULT_LIST_CONFIG: IListConfig = {
  requestListTitle: 'Group Membership Requests',
  groupListTitle: 'Managed Groups',
  authorizedAdminsListTitle: 'Group Management Authorized Admins',
  sitePermissionsListTitle: 'Group Site Permissions'
};

export class AccountManagementService {
  private _context: WebPartContext;
  private _webUrl: string;
  private _config: IListConfig;

  public constructor(context: WebPartContext, config?: Partial<IListConfig>) {
    this._context = context;
    this._config = {
      requestListTitle: (config && config.requestListTitle) || DEFAULT_LIST_CONFIG.requestListTitle,
      groupListTitle: (config && config.groupListTitle) || DEFAULT_LIST_CONFIG.groupListTitle,
      authorizedAdminsListTitle:
        (config && config.authorizedAdminsListTitle) || DEFAULT_LIST_CONFIG.authorizedAdminsListTitle,
      sitePermissionsListTitle:
        (config && config.sitePermissionsListTitle) || DEFAULT_LIST_CONFIG.sitePermissionsListTitle,
      listSiteUrl: (config && config.listSiteUrl) || ''
    };
    // The four lists can live on ONE central site so the web part works on any page in the site
    // collection. Honor an explicit List site URL only when it is the same tenant (same host) as the
    // page; otherwise fall back to the current site.
    this._webUrl = AccountManagementService._sameHostOrDefault(
      this._config.listSiteUrl,
      context.pageContext.web.absoluteUrl
    );
  }

  /** Return `candidate` (trimmed) when it is a same-host URL as `fallback`; otherwise `fallback`. */
  private static _sameHostOrDefault(candidate: string | undefined, fallback: string): string {
    const base: string = fallback.replace(/\/+$/, '');
    if (!candidate) {
      return base;
    }
    const site: string = candidate.replace(/\/+$/, '');
    try {
      return new URL(site).host.toLowerCase() === new URL(base).host.toLowerCase() ? site : base;
    } catch {
      return base;
    }
  }

  public async getAuthorizedGroups(): Promise<IOfficeGroup[]> {
    diag(`${DIAG} SharePoint current user request starting`);
    const current: ICurrentUser = await this._getCurrentUser();
    diag(`${DIAG} SharePoint current user loaded`, {
      userId: current.Id,
      userPrincipalName: current.UserPrincipalName || current.Email || current.Title
    });

    const select: string = [
      'Id',
      'OfficeGroupRecord/Id',
      'OfficeGroupRecord/Title',
      'OfficeGroupRecord/GroupId',
      'OfficeGroupRecord/SiteTitle'
    ].join(',');
    const url: string =
      `${this._webUrl}/_api/web/lists/getbytitle('${this._config.authorizedAdminsListTitle}')/items` +
      `?$select=${select}&$expand=OfficeGroupRecord&$filter=UserId eq ${current.Id}&$top=5000`;
    diag(`${DIAG} authorized admins request starting`, {
      listTitle: this._config.authorizedAdminsListTitle,
      currentUserId: current.Id
    });
    const items: any[] = (await this._get(url)).value;
    diag(`${DIAG} authorized admins loaded`, { count: items.length });

    // OfficeGroupRecord can be a single-value lookup (object) OR a multi-value lookup (array / {results:[]}),
    // so ONE admin row may authorize several groups. Normalize each row's lookup to a flat list of records.
    const asRecords = (v: any): any[] =>
      !v ? [] : Array.isArray(v) ? v : Array.isArray(v.results) ? v.results : [v];

    const recordIds: number[] = [];
    items.forEach((i: any) => {
      asRecords(i.OfficeGroupRecord).forEach((r: any) => {
        if (r && typeof r.Id === 'number') {
          recordIds.push(r.Id);
        }
      });
    });
    if (recordIds.length === 0) {
      return [];
    }

    diag(`${DIAG} group records request starting`, {
      count: recordIds.length,
      listTitle: this._config.groupListTitle
    });
    const groupsById: Map<number, IOfficeGroup> = await this._getOfficeGroupsByIds(
      Array.from(new Set<number>(recordIds))
    );
    diag(`${DIAG} group records loaded`, { count: groupsById.size });

    // Flatten every (row -> its group records) into a deduped list of authorized groups.
    const authorized: IOfficeGroup[] = [];
    const seen: Set<number> = new Set<number>();
    items.forEach((i: any) => {
      asRecords(i.OfficeGroupRecord).forEach((r: any) => {
        const recId: number | undefined = r && typeof r.Id === 'number' ? r.Id : undefined;
        const group: IOfficeGroup | undefined = recId !== undefined ? groupsById.get(recId) : undefined;
        if (group && !seen.has(group.id)) {
          seen.add(group.id);
          authorized.push({ ...group, authorizationItemId: i.Id });
        }
      });
    });
    return authorized;
  }

  /** PROD path: file a Pending request item for the backend flow to process. */
  public async createMembershipRequest(input: IMembershipChangeInput): Promise<IMembershipRequest> {
    if (!input.group.groupId) {
      throw new Error('The selected group record is missing the Microsoft 365 GroupId value.');
    }
    if (!input.member.id) {
      throw new Error('The selected user is missing the Entra ID value.');
    }
    const now: string = new Date().toISOString();
    const payload: any = {
      Title: `${input.action}: ${input.member.displayName}`,
      Action: input.action,
      GroupId: input.group.groupId,
      GroupName: input.group.title,
      MemberId: input.member.id,
      MemberDisplayName: input.member.displayName,
      Status: 'Pending',
      RequestedOn: now,
      TargetUserPrincipalName: input.member.userPrincipalName || input.member.mail || '',
      TargetUserEmail: input.member.mail || input.member.userPrincipalName || '',
      CorrelationId: Guid.newGuid().toString(),
      OfficeGroupRecordId: input.group.id
    };
    if (input.justification && input.justification.trim()) {
      payload.Justification = input.justification.trim();
    }
    payload.MemberEntraIdId = await this._ensureUserId(input.member.userPrincipalName || input.member.mail);
    payload.RequestedById = await this._currentUserId(); // Person column "RequestedBy" = the signed-in initiator

    return this._mapRequest(await this._postRequestItem(payload));
  }

  /**
   * Audit record for a SharePoint-group change that was applied directly (no flow).
   * Written as Completed so the flow skips it (not Pending); best-effort so it never blocks the change.
   */
  public async recordCompletedChange(input: IMembershipChangeInput): Promise<boolean> {
    try {
      const now: string = new Date().toISOString();
      const payload: any = {
        Title: `${input.action}: ${input.member.displayName}`,
        Action: input.action,
        GroupId: input.group.groupId,
        GroupName: input.group.title,
        MemberId: input.member.id,
        MemberDisplayName: input.member.displayName,
        Status: 'Completed',
        ResultMessage: 'Applied directly to the SharePoint group.',
        RequestedOn: now,
        TargetUserPrincipalName: input.member.userPrincipalName || input.member.mail || '',
        TargetUserEmail: input.member.mail || input.member.userPrincipalName || '',
        CorrelationId: Guid.newGuid().toString(),
        OfficeGroupRecordId: input.group.id
      };
      if (input.justification && input.justification.trim()) {
        payload.Justification = input.justification.trim();
      }
      payload.RequestedById = await this._currentUserId(); // Person column "RequestedBy" = the signed-in initiator
      await this._postRequestItem(payload);
      return true;
    } catch (e) {
      console.warn('365 Account Management: could not write the SharePoint-change audit record.', e);
      return false;
    }
  }

  /** POST a request-list item; if the list has no Justification column yet, retry without it. */
  private async _postRequestItem(payload: any): Promise<any> {
    const url: string = `${this._webUrl}/_api/web/lists/getbytitle('${this._config.requestListTitle}')/items`;
    try {
      return await this._post(url, payload);
    } catch (e) {
      // RequestedBy is optional metadata; if the list lacks that column, retry without it. We do NOT
      // strip Justification — a required reason must be durably stored, so a missing Justification column
      // surfaces as an error rather than silently dropping the reason.
      const optional: string[] = ['RequestedById'];
      const present: string[] = optional.filter((k: string) => Object.prototype.hasOwnProperty.call(payload, k));
      if (present.length > 0) {
        const fallback: any = { ...payload };
        present.forEach((k: string) => delete fallback[k]);
        console.warn(`365 Account Management: request list missing optional column(s) ${present.join(', ')}; submitting without them.`);
        return this._post(url, fallback);
      }
      throw e;
    }
  }

  /**
   * DEV path: change SharePoint site-group membership directly in the caller's context (no flow).
   * Add uses the target UPN as a claims LoginName; remove uses the SharePoint site-user id.
   */
  public async changeSharePointGroupMembership(input: {
    action: MembershipAction;
    spGroupId: string;
    member: IUser;
    siteUrl?: string;
  }): Promise<void> {
    const gid: string = encodeURIComponent(input.spGroupId);
    const web: string = this._sameTenantOrPageWeb(input.siteUrl); // SP groups may live on another (same-tenant) site

    if (input.action === 'Add Member') {
      // Resolve the principal via ensureuser, then add the canonical login to the site group.
      // - User: resolve by UPN/mail; fall back to the i:0#.f|membership| claim.
      // - Group: resolve by mail when present (M365 + mail-enabled groups) else by the Azure AD
      //   security-group tenant claim c:0t.c|tenant|<objectId>; fall back to that same candidate.
      let candidate: string;
      let fallback: string;
      if (input.member.isGroup) {
        const byMailOrClaim: string | undefined =
          input.member.mail || (input.member.id ? `c:0t.c|tenant|${input.member.id}` : undefined);
        if (!byMailOrClaim) {
          throw new Error('The selected group cannot be resolved (no mail or id).');
        }
        candidate = byMailOrClaim;
        fallback = byMailOrClaim;
      } else {
        const login: string | undefined = input.member.userPrincipalName || input.member.mail;
        if (!login) {
          throw new Error('The selected user has no UPN/email to add.');
        }
        candidate = login;
        fallback = `i:0#.f|membership|${login}`;
      }
      let loginName: string = fallback;
      try {
        const ensured: any = await this._post(`${web}/_api/web/ensureuser`, { logonName: candidate });
        if (ensured && ensured.LoginName) {
          loginName = ensured.LoginName;
        }
      } catch {
        /* fall back to the constructed claim */
      }
      diag(`${DIAG} SharePoint group add`, { spGroupId: input.spGroupId, isGroup: !!input.member.isGroup, loginName });
      await this._postSiteGroup(
        `${web}/_api/web/sitegroups(${gid})/users`,
        { __metadata: { type: 'SP.User' }, LoginName: loginName },
        'Add to SharePoint group'
      );
    } else {
      diag(`${DIAG} SharePoint group remove`, { spGroupId: input.spGroupId });
      await this._postSiteGroup(
        `${web}/_api/web/sitegroups(${gid})/users/removebyid(${encodeURIComponent(input.member.id)})`,
        undefined,
        'Remove from SharePoint group'
      );
    }
  }

  /** POST to a sitegroups endpoint, surfacing the real SharePoint error body on failure. */
  private async _postSiteGroup(url: string, verboseBody: any | undefined, label: string): Promise<void> {
    const options: any = {
      headers: verboseBody
        ? {
            // SPHttpClient defaults to OData v4 (JSON-light); the SP.User entity POST needs OData v3 (verbose).
            // Capital "Content-Type" + empty "odata-version" forces verbose, else SharePoint returns
            // "Parsing JSON Light feeds or entries in requests without entity set is not supported".
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': ''
          }
        : { Accept: 'application/json;odata=nometadata' }
    };
    if (verboseBody) {
      options.body = JSON.stringify(verboseBody);
    }
    const resp: SPHttpClientResponse = await this._context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );
    if (!resp.ok) {
      const text: string = await resp.text();
      const detail: string = toMessage(text, '');
      console.error('365 Account Management SharePoint group REST failed', {
        url,
        status: resp.status,
        responseText: text
      });
      throw new Error(`${label} failed (HTTP ${resp.status})${detail ? ': ' + detail : ''}.`);
    }
  }

  public async getRequest(id: number): Promise<IMembershipRequest> {
    const url: string =
      `${this._webUrl}/_api/web/lists/getbytitle('${this._config.requestListTitle}')/items(${id})` +
      `?$select=Id,Status,ResultMessage,AuthorizationChecked,AuthorizationResult`;
    return this._mapRequest(await this._get(url));
  }

  public async pollRequest(
    id: number,
    onUpdate: (request: IMembershipRequest) => void,
    timeoutMs: number = 120000
  ): Promise<IMembershipRequest> {
    const start: number = Date.now();
    let attempt: number = 0;
    while (Date.now() - start < timeoutMs) {
      const request: IMembershipRequest = await this.getRequest(id);
      onUpdate(request);
      if (this.isTerminalStatus(request.status)) {
        return request;
      }
      // Backoff: responsive early (2.5s), easing to 8s — far fewer calls for slow flows than a flat 3s.
      const wait: number = Math.min(2500 + attempt * 1000, 8000);
      attempt++;
      await this._delay(wait);
    }
    throw new Error('The request was submitted, but the flow did not finish before the timeout.');
  }

  public isTerminalStatus(status: string): boolean {
    return ['Completed', 'Failed'].indexOf(status) !== -1;
  }

  private async _getOfficeGroupsByIds(ids: number[]): Promise<Map<number, IOfficeGroup>> {
    const unique: number[] = ids.filter((v: number, i: number, a: number[]) => a.indexOf(v) === i);
    const baseCols: string[] = [
      'Id',
      'Title',
      'GroupId',
      'Description',
      'Mail',
      'MailNickname',
      'Visibility',
      'CreatedDateTime',
      'IsTeamsConnected',
      'SiteUrl',
      'SiteTitle'
    ];
    const map: Map<number, IOfficeGroup> = new Map<number, IOfficeGroup>();
    const batch: number = 40;
    // GroupType is an OPTIONAL column; select it when present, else fall back so existing lists still load.
    let includeGroupType: boolean = true;
    for (let i: number = 0; i < unique.length; i += batch) {
      const chunk: number[] = unique.slice(i, i + batch);
      const filter: string = chunk.map((id: number) => `Id eq ${id}`).join(' or ');
      const listUrl: string = `${this._webUrl}/_api/web/lists/getbytitle('${this._config.groupListTitle}')/items`;
      const cols: string[] = includeGroupType ? baseCols.concat(['GroupType']) : baseCols;
      let res: any;
      try {
        res = await this._get(`${listUrl}?$select=${cols.join(',')}&$filter=${filter}&$top=5000`);
      } catch (e) {
        if (includeGroupType) {
          includeGroupType = false; // column absent — retry this and later batches without it
          res = await this._get(`${listUrl}?$select=${baseCols.join(',')}&$filter=${filter}&$top=5000`);
        } else {
          throw e;
        }
      }
      res.value.forEach((e: any) => map.set(e.Id, this._mapOfficeGroup(e)));
    }
    return map;
  }

  private async _getCurrentUser(): Promise<ICurrentUser> {
    return this._get(`${this._webUrl}/_api/web/currentuser?$select=Id,Title,Email,UserPrincipalName`);
  }

  /** Cached SharePoint site-user id of the signed-in user (for the RequestedBy Person column). */
  private _currentUserIdCache?: number;
  private async _currentUserId(): Promise<number | undefined> {
    if (this._currentUserIdCache === undefined) {
      try {
        const u: ICurrentUser = await this._getCurrentUser();
        this._currentUserIdCache = u && typeof u.Id === 'number' ? u.Id : undefined;
      } catch {
        return undefined;
      }
    }
    return this._currentUserIdCache;
  }

  private async _ensureUserId(login: string | undefined): Promise<number | undefined> {
    if (!login) {
      return undefined;
    }
    try {
      const result: any = await this._post(`${this._webUrl}/_api/web/ensureuser`, { logonName: login });
      return result.Id;
    } catch {
      return undefined;
    }
  }

  private _mapOfficeGroup(e: any): IOfficeGroup {
    const siteUrl: string | undefined = typeof e.SiteUrl === 'string' ? e.SiteUrl : e.SiteUrl && e.SiteUrl.Url;
    return {
      id: e.Id,
      title: e.Title || e.SiteTitle || e.MailNickname || 'Untitled group',
      groupId: e.GroupId || '',
      description: e.Description,
      mail: e.Mail,
      mailNickname: e.MailNickname,
      visibility: e.Visibility,
      createdDateTime: e.CreatedDateTime,
      isTeamsConnected: e.IsTeamsConnected,
      siteUrl: siteUrl,
      siteTitle: e.SiteTitle,
      groupType: e.GroupType
    };
  }

  /** The current user's most recent requests (works under item-level "read own" hardening). */
  public async getRecentRequests(top: number = 10): Promise<IRequestSummary[]> {
    const current: ICurrentUser = await this._getCurrentUser();
    const select: string = [
      'Id',
      'Action',
      'GroupName',
      'GroupId',
      'MemberDisplayName',
      'Status',
      'ResultMessage',
      'RequestedOn',
      'Modified',
      'Author/Title'
    ].join(',');
    const url: string =
      `${this._webUrl}/_api/web/lists/getbytitle('${this._config.requestListTitle}')/items` +
      `?$select=${select}&$expand=Author&$filter=AuthorId eq ${current.Id}&$orderby=Id desc&$top=${top}`;
    const items: any[] = (await this._get(url)).value || [];
    return items.map((e: any): IRequestSummary => ({
      id: e.Id,
      action: e.Action,
      groupName: e.GroupName,
      groupId: e.GroupId,
      memberDisplayName: e.MemberDisplayName,
      status: e.Status || 'Pending',
      resultMessage: this._optional(e.ResultMessage),
      requestedOn: e.RequestedOn,
      requesterName: e.Author && e.Author.Title,
      modified: e.Modified
    }));
  }

  /** Curated 'Group Site Permissions' rows for a group (the sites it's used on + permission level). */
  public async getGroupSitePermissions(groupId: string): Promise<ISitePermission[]> {
    if (!groupId) {
      return [];
    }
    const select: string = ['Id', 'SiteName', 'SiteUrl', 'Permission'].join(',');
    const url: string =
      `${this._webUrl}/_api/web/lists/getbytitle('${this._config.sitePermissionsListTitle}')/items` +
      `?$select=${select}&$filter=GroupId eq '${encodeURIComponent(groupId)}'&$orderby=SiteName&$top=500`;
    try {
      const items: any[] = (await this._get(url)).value || [];
      return items.map((e: any): ISitePermission => ({
        siteName: e.SiteName || (e.SiteUrl && e.SiteUrl.Description) || '(site)',
        siteUrl: (e.SiteUrl && e.SiteUrl.Url) || (typeof e.SiteUrl === 'string' ? e.SiteUrl : undefined),
        permission: e.Permission || ''
      }));
    } catch {
      return []; // optional feature — list may not be provisioned yet, or no read access
    }
  }

  private _mapRequest(e: any): IMembershipRequest {
    return {
      id: e.Id,
      status: e.Status || 'Pending',
      resultMessage: this._optional(e.ResultMessage),
      authorizationChecked: e.AuthorizationChecked,
      authorizationResult: this._optional(e.AuthorizationResult)
    };
  }

  private _optional(value: any): string | undefined {
    if (!value) {
      return undefined;
    }
    const message: string = toMessage(value, '');
    return message || undefined;
  }

  private async _get(url: string): Promise<any> {
    diag(`${DIAG} SharePoint GET`, { url });
    const resp: SPHttpClientResponse = await this._context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    return this._parseResponse(resp);
  }

  private async _post(url: string, body: any): Promise<any> {
    diag(`${DIAG} SharePoint POST`, { url, payloadKeys: Object.keys(body) });
    const resp: SPHttpClientResponse = await this._context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(body)
      }
    );
    return this._parseResponse(resp);
  }

  private async _parseResponse(resp: SPHttpClientResponse): Promise<any> {
    const text: string = await resp.text();
    if (!resp.ok) {
      let message: string =
        `SharePoint request failed with status ${resp.status}` +
        (resp.statusText ? ` (${resp.statusText})` : '') +
        '.';
      message = toMessage(text, message);
      console.error('365 Account Management SharePoint REST request failed', {
        url: resp.url,
        status: resp.status,
        statusText: resp.statusText,
        message: message,
        responseText: text
      });
      throw new Error(message);
    }
    return text ? JSON.parse(text) : undefined;
  }

  /** Use the override site only when it's the same tenant (same host) as the page; else the page web. */
  private _sameTenantOrPageWeb(siteUrl: string | undefined): string {
    const pageWeb: string = this._webUrl.replace(/\/+$/, '');
    if (!siteUrl) {
      return pageWeb;
    }
    const site: string = siteUrl.replace(/\/+$/, '');
    try {
      if (new URL(site).host.toLowerCase() === new URL(pageWeb).host.toLowerCase()) {
        return site;
      }
      diag(`${DIAG} group SiteUrl is a different host than the page; ignoring it and using the page web`, { siteUrl });
      return pageWeb;
    } catch {
      diag(`${DIAG} group SiteUrl could not be parsed; using the page web`, { siteUrl });
      return pageWeb;
    }
  }

  private _delay(ms: number): Promise<void> {
    return new Promise<void>((resolve: () => void) => window.setTimeout(resolve, ms));
  }
}
