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
  public async getGroupMembers(groupId: string): Promise<IUser[]> {
    if (!groupId) {
      throw new Error('The selected group record is missing the Microsoft 365 GroupId value.');
    }

    if (isSharePointGroup(groupId)) {
      return this._getSharePointGroupMembers(groupId);
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
  private async _getSharePointGroupMembers(spGroupId: string): Promise<IUser[]> {
    const web: string = this._context.pageContext.web.absoluteUrl;
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
      throw new Error(`Unable to load SharePoint group members (HTTP ${resp.status}).`);
    }
    const data: any = await resp.json();
    const value: any[] = (data && data.value) || [];
    return value
      .filter((u: any) => u.PrincipalType === 1) // 1 = User; drop nested SharePoint/security groups
      .map((u: any): IUser => ({
        id: String(u.Id), // SharePoint site-user Id -> used by removebyid on the DEV remove path
        displayName: u.Title || u.UserPrincipalName || u.Email || 'Unknown user',
        mail: u.Email,
        userPrincipalName: u.UserPrincipalName || u.LoginName
      }))
      .filter((u: IUser) => !!u.displayName)
      .sort(byDisplayName);
  }

  /** Directory search (tenant-wide; works in DEV and PROD regardless of group type). */
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

  private _mapUser(raw: any): IUser {
    return {
      id: raw.id,
      displayName: raw.displayName || raw.userPrincipalName || raw.mail || 'Unknown user',
      mail: raw.mail,
      userPrincipalName: raw.userPrincipalName,
      jobTitle: raw.jobTitle
    };
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
