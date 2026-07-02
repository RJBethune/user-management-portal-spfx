import { diag } from '../shared/log';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IUser } from '../models/types';
import { isSharePointGroup } from '../shared/groupType';
import { toMessage } from '../shared/errors';

/* eslint-disable @typescript-eslint/no-explicit-any */

const DIAG: string = '365 Account Management diagnostic';

/** Locale-aware A→Z by display name (used for member and owner lists, NOT search results). */
const byDisplayName = (a: IUser, b: IUser): number =>
  (a.displayName || '').localeCompare(b.displayName || '', undefined, { sensitivity: 'base' });

export class GraphService {
  private _context: WebPartContext;

  public constructor(context: WebPartContext) {
    this._context = context;
  }

  /** Members of a group. O365 group (GUID) -> Graph; SharePoint site group (int) -> SharePoint REST. */
  public async getGroupMembers(groupId: string, siteUrl?: string): Promise<IUser[]> {
    if (!groupId) {
      throw new Error('The selected group record is missing the Microsoft 365 GroupId value.');
    }

    if (isSharePointGroup(groupId)) {
      return this._getSharePointGroupMembers(groupId, siteUrl);
    }

    try {
      diag(`${DIAG} Graph group members request starting`, { groupId });
      const client: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      const collected: any[] = [];
      let path: string | undefined =
        `/groups/${groupId}/members/microsoft.graph.user` +
        `?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=999`;
      while (path) {
        const page: any = await client.api(this._toGraphApiPath(path)).get();
        collected.push(...page.value);
        path = page['@odata.nextLink'];
      }
      diag(`${DIAG} Graph group members loaded`, { groupId, count: collected.length });
      return collected.map(this._mapUser).filter((u: IUser) => !!u.displayName).sort(byDisplayName);
    } catch (err) {
      throw new Error('Unable to load group members. ' + toMessage(err, 'Microsoft Graph request failed.'));
    }
  }

  /** DEV mock: read members of a SharePoint site group via REST, mapped to the IUser shape. */
  private async _getSharePointGroupMembers(spGroupId: string, siteUrl?: string): Promise<IUser[]> {
    const web: string = this._sameTenantWeb(siteUrl);
    const url: string =
      `${web}/_api/web/sitegroups(${encodeURIComponent(spGroupId)})/users` +
      `?$select=Id,Title,Email,LoginName,UserPrincipalName,PrincipalType&$top=999`;
    diag(`${DIAG} SharePoint group members request starting`, { spGroupId });
    const resp: SPHttpClientResponse = await this._context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!resp.ok) {
      const detail: string = await this._readSpError(resp);
      throw new Error(
        `Couldn't load this group's members — ${detail} ` +
          `Check that the group's Site URL (${web}) is a site that exists, contains site group ${spGroupId}, and that you have access to it.`
      );
    }
    const data: any = await resp.json();
    const value: any[] = (data && data.value) || [];
    return value
      // Keep users AND nested group principals so nested groups are visible, not silently hidden.
      // PrincipalType: 1=User, 2=Distribution list, 4=Security/O365 group, 8=SharePoint group.
      .map((u: any): IUser => ({
        id: String(u.Id), // SharePoint site-user Id -> used by removebyid on the remove path
        displayName: u.Title || u.UserPrincipalName || u.Email || 'Unknown principal',
        mail: u.Email,
        userPrincipalName: u.UserPrincipalName || u.LoginName,
        isGroup: u.PrincipalType !== 1
      }))
      .filter((u: IUser) => !!u.displayName)
      .sort(byDisplayName);
  }

  /** Extract SharePoint's own error text (status + message) from a failed REST response. */
  private async _readSpError(resp: SPHttpClientResponse): Promise<string> {
    let message: string = '';
    try {
      const body: any = await resp.json();
      const m: any = body && body.error && body.error.message;
      message = (m && typeof m === 'object' ? m.value : m) || '';
    } catch {
      message = '';
    }
    return `HTTP ${resp.status}${message ? ': ' + message : ''}.`;
  }

  /**
   * Directory search across the whole tenant (Graph /users?$search). By design this surfaces ANY directory
   * user to an authorized admin (so they can be added to a group). This is a deliberate, tenant-wide
   * data-exposure decision. Works in DEV and PROD regardless of group type.
   */
  public async searchUsers(query: string): Promise<IUser[]> {
    const term: string = (query || '').trim();
    if (term.length < 2) {
      return [];
    }
    try {
      diag(`${DIAG} Graph user search starting`, { queryLength: term.length });
      const client: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      const escaped: string = term.replace(/"/g, '\\"');
      const search: string =
        `"displayName:${escaped}" OR "mail:${escaped}" OR "userPrincipalName:${escaped}"`;
      const result: any = await client
        .api(
          `/users?$search=${encodeURIComponent(search)}` +
            `&$select=id,displayName,mail,userPrincipalName,jobTitle&$top=25`
        )
        .header('ConsistencyLevel', 'eventual')
        .get();
      diag(`${DIAG} Graph user search loaded`, { count: result.value.length });
      return result.value.map(this._mapUser).filter((u: IUser) => !!u.displayName);
    } catch (err) {
      throw new Error('Unable to search users. ' + toMessage(err, 'Microsoft Graph request failed.'));
    }
  }

  /** Read-only owners of an O365 group (GUID only). Best-effort; throws on failure for the caller. */
  public async getGroupOwners(groupId: string): Promise<IUser[]> {
    if (!groupId || isSharePointGroup(groupId)) {
      return [];
    }
    diag(`${DIAG} Graph group owners request starting`, { groupId });
    const client: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
    const result: any = await client
      .api(
        `/groups/${groupId}/owners/microsoft.graph.user` +
          `?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=100`
      )
      .get();
    return (((result && result.value) || []) as any[])
      .map(this._mapUser)
      .filter((u: IUser) => !!u.displayName)
      .sort(byDisplayName);
  }

  /** Object-URL of an O365 group's photo, or undefined when none/404 or not an O365 group. */
  public async getGroupPhotoUrl(groupId: string): Promise<string | undefined> {
    if (!groupId || isSharePointGroup(groupId)) {
      return undefined;
    }
    try {
      const client: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      const blob: Blob = await client.api(`/groups/${groupId}/photo/$value`).responseType('blob' as any).get();
      return blob && blob.size > 0 ? URL.createObjectURL(blob) : undefined;
    } catch {
      return undefined; // 404 (no custom photo) or any error — caller shows the initials tile
    }
  }

  private _mapUser(raw: any): IUser {
    return {
      id: raw.id,
      displayName: raw.displayName || raw.userPrincipalName || raw.mail || 'Unknown user',
      mail: raw.mail,
      userPrincipalName: raw.userPrincipalName,
      jobTitle: raw.jobTitle
    };
  }

  /** Use the override site only when it's the same tenant (same host) as the page; else the page web. */
  private _sameTenantWeb(siteUrl: string | undefined): string {
    const pageWeb: string = this._context.pageContext.web.absoluteUrl.replace(/\/+$/, '');
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

  /** Strip the absolute host + version segment from an @odata.nextLink so it can be re-fed to .api(). */
  private _toGraphApiPath(value: string): string {
    if (value.indexOf('https://') !== 0) {
      return value;
    }
    try {
      const url: URL = new URL(value);
      const path: string = url.pathname.replace(/^\/(v1\.0|beta)/i, '');
      return `${path}${url.search}`;
    } catch {
      return value;
    }
  }
}
