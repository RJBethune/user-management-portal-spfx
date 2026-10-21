import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import { IListConfig } from './AccountManagementService';
import { IAccessAssignment, IAccessOverview, IManagedGroupReference, lookupIds, summarizeAssignments } from '../models/accessAssignments';

interface IListMetadata {
  Id: string;
  DefaultViewUrl: string;
  DefaultEditFormUrl: string;
  DefaultNewFormUrl: string;
  EffectiveBasePermissions: { High: string; Low: string };
}
interface IAssignmentRow {
  Id: number;
  User?: { Id: number; Title: string; EMail?: string };
  OfficeGroupRecord?: unknown;
}
interface IGroupRow { Id: number; Title: string; GroupId: string; IsTeamsConnected?: boolean; }

/** Read-only administration. SharePoint owns edit forms and enforces all list writes. */
export class AccessAdministrationService {
  private readonly webUrl: string;
  public constructor(private readonly context: WebPartContext, private readonly config: IListConfig) {
    const fallback = context.pageContext.web.absoluteUrl.replace(/\/+$/, '');
    try {
      const candidate = new URL((config.listSiteUrl || fallback).trim());
      this.webUrl = candidate.origin === new URL(fallback).origin && candidate.protocol === 'https:'
        ? candidate.href.replace(/\/+$/, '') : fallback;
    } catch { this.webUrl = fallback; }
  }

  private listApi(title: string): string {
    return `${this.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title.replace(/'/g, "''"))}')`;
  }

  private async get<T>(url: string): Promise<T> {
    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata' }
    });
    if (!response.ok) throw new Error(`Access assignments could not be read (HTTP ${response.status}). Check list permissions and try again.`);
    return response.json();
  }

  private async metadata(): Promise<IListMetadata> {
    return this.get<IListMetadata>(`${this.listApi(this.config.authorizedAdminsListTitle)}?$select=Id,DefaultViewUrl,DefaultEditFormUrl,DefaultNewFormUrl,EffectiveBasePermissions`);
  }

  private canManage(metadata: IListMetadata): boolean {
    const permissions = new SPPermission({
      High: Number(metadata.EffectiveBasePermissions.High),
      Low: Number(metadata.EffectiveBasePermissions.Low)
    });
    return permissions.hasPermission(SPPermission.manageLists) && permissions.hasPermission(SPPermission.editListItems);
  }

  public async isAvailable(): Promise<boolean> {
    try { return this.canManage(await this.metadata()); } catch { return false; }
  }

  private async all<T>(url: string): Promise<T[]> {
    const rows: T[] = [];
    const visited = new Set<string>();
    let next: string | undefined = url;
    while (next) {
      const target: URL = new URL(next, this.webUrl);
      if (!target.href.startsWith(`${this.webUrl}/_api/`) || visited.has(target.href)) throw new Error('Invalid assignment list continuation. Refresh and try again.');
      visited.add(target.href);
      const page: { value: T[]; '@odata.nextLink'?: string; 'odata.nextLink'?: string } = await this.get(target.href);
      if (!Array.isArray(page.value)) throw new Error('The assignment list returned an unexpected response.');
      rows.push(...page.value);
      next = page['@odata.nextLink'] || page['odata.nextLink'];
    }
    return rows;
  }

  private formUrl(value: string): string {
    if (typeof value !== 'string' || !value.trim()) throw new Error('The assignment list did not return a form URL.');
    const url = new URL(value, this.webUrl);
    if (url.origin !== new URL(this.webUrl).origin || url.protocol !== 'https:') throw new Error('The assignment list returned an invalid form URL.');
    return url.href;
  }

  public async load(): Promise<IAccessOverview> {
    const metadata = await this.metadata();
    if (!this.canManage(metadata)) throw new Error('Managing access assignments requires Manage Lists and Edit Items on the Authorized Admins list.');
    const [rawAssignments, rawGroups] = await Promise.all([
      this.all<IAssignmentRow>(`${this.listApi(this.config.authorizedAdminsListTitle)}/items?$select=Id,User/Id,User/Title,User/EMail,OfficeGroupRecord/Id&$expand=User,OfficeGroupRecord&$top=500`),
      this.all<IGroupRow>(`${this.listApi(this.config.groupListTitle)}/items?$select=Id,Title,GroupId,IsTeamsConnected&$top=500`)
    ]);
    const groups = rawGroups.map((row): IManagedGroupReference => ({ id: row.Id, title: row.Title || row.GroupId || `Group record ${row.Id}`, groupId: row.GroupId || '', isTeamsConnected: !!row.IsTeamsConnected }))
      .sort((a, b) => a.title.localeCompare(b.title));
    const byId = new Map(groups.map((group): [number, IManagedGroupReference] => [group.id, group]));
    const assignments = rawAssignments.map((row): IAccessAssignment => {
      const ids = lookupIds(row.OfficeGroupRecord);
      return {
        id: row.Id, userId: row.User && row.User.Id,
        userName: row.User && row.User.Title || 'Missing admin', userEmail: row.User && row.User.EMail || '',
        groups: ids.filter((id) => byId.has(id)).map((id) => byId.get(id) as IManagedGroupReference).sort((a, b) => a.title.localeCompare(b.title)),
        missingGroupIds: ids.filter((id) => !byId.has(id))
      };
    }).sort((a, b) => a.userName.localeCompare(b.userName));
    const editUrl = this.formUrl(metadata.DefaultEditFormUrl);
    return {
      assignments, groups, ...summarizeAssignments(assignments, groups),
      listUrl: this.formUrl(metadata.DefaultViewUrl), newItemUrl: this.formUrl(metadata.DefaultNewFormUrl),
      editItemUrl: (id) => { const url = new URL(editUrl); url.searchParams.set('ID', String(id)); return url.href; }
    };
  }
}
