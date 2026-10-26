import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import { IListConfig } from './AccountManagementService';
import { GraphService } from './GraphService';
import { IUser } from '../models/types';
import { DEFAULT_ACCESS_ADMIN_GROUP, IAccessAssignment, IAccessOverview, IAssignmentEditor, IManagedGroupReference, lookupIds, summarizeAssignments } from '../models/accessAssignments';

interface IListMetadata {
  Id: string;
  DefaultViewUrl: string;
  ListItemEntityTypeFullName: string;
  EffectiveBasePermissions: { High: string; Low: string };
}
interface IAssignmentRow {
  Id: number;
  User?: { Id: number; Title: string; EMail?: string };
  OfficeGroupRecord?: unknown;
}
interface IGroupRow { Id: number; Title: string; GroupId: string; IsTeamsConnected?: boolean; }
interface IAssignmentField { TypeAsString: string; ReadOnlyField: boolean; Required: boolean; LookupList?: string; SelectionGroup?: number; }
interface IEditorSchema { metadata: IListMetadata; user: IAssignmentField; lookup: IAssignmentField; }

export class AssignmentSaveError extends Error {
  public constructor(message: string, public readonly code: 'conflict' | 'denied' | 'uncertain' | 'validation') {
    super(message);
    Object.setPrototypeOf(this, AssignmentSaveError.prototype);
  }
}

/** Reads and writes the existing list as the signed-in user. SharePoint remains the authorization boundary. */
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
    return this.get<IListMetadata>(`${this.listApi(this.config.authorizedAdminsListTitle)}?$select=Id,DefaultViewUrl,ListItemEntityTypeFullName,EffectiveBasePermissions`);
  }

  private canManage(metadata: IListMetadata): boolean {
    return this.hasPermission(metadata, SPPermission.manageLists) && this.hasPermission(metadata, SPPermission.editListItems);
  }

  private hasPermission(metadata: IListMetadata, permission: SPPermission): boolean {
    const permissions = new SPPermission({
      High: Number(metadata.EffectiveBasePermissions.High),
      Low: Number(metadata.EffectiveBasePermissions.Low)
    });
    return permissions.hasPermission(permission);
  }

  public async isAvailable(): Promise<boolean> {
    try { await this.requireAdminGroup(); return this.canManage(await this.metadata()); } catch { return false; }
  }

  private async requireAdminGroup(): Promise<void> {
    const name = (this.config.accessAdminGroupName === undefined ? DEFAULT_ACCESS_ADMIN_GROUP : this.config.accessAdminGroupName).trim();
    if (!name) throw new AssignmentSaveError('Access assignments is disabled until a SharePoint admin group is configured.', 'denied');
    // Check the signed-in user's direct SharePoint group membership on the list site, never the page's site.
    // Do not cache this decision: a removed member must not save through an already-open editor.
    const filter = encodeURIComponent(`Title eq '${name.replace(/'/g, "''")}'`);
    let groups: { Id: number; Title: string }[];
    try { groups = await this.all(`${this.webUrl}/_api/web/currentuser/groups?$select=Id,Title&$filter=${filter}`); }
    catch { throw new AssignmentSaveError('Your access administration group membership could not be verified. Try again or contact a site owner.', 'denied'); }
    if (!groups.some(group => Number.isSafeInteger(group.Id) && group.Id > 0 && typeof group.Title === 'string' && group.Title.toLowerCase() === name.toLowerCase())) {
      throw new AssignmentSaveError(`Access assignments requires membership in the SharePoint group ${name}. Your draft has not been saved.`, 'denied');
    }
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

  private async groups(): Promise<IManagedGroupReference[]> {
    const rows = await this.all<IGroupRow>(`${this.listApi(this.config.groupListTitle)}/items?$select=Id,Title,GroupId,IsTeamsConnected&$top=500`);
    return rows.map((row) => ({ id: row.Id, title: row.Title || row.GroupId || `Group record ${row.Id}`, groupId: row.GroupId || '', isTeamsConnected: !!row.IsTeamsConnected }))
      .sort((a, b) => a.title.localeCompare(b.title));
  }

  private async editorSchema(creating: boolean): Promise<IEditorSchema> {
    // Recheck on each open and save. A cached UI gate never authorizes a write.
    await this.requireAdminGroup();
    const metadata = await this.metadata();
    if (!this.canManage(metadata) || (creating && !this.hasPermission(metadata, SPPermission.addListItems))) {
      throw new AssignmentSaveError(`Your permissions no longer allow ${creating ? 'adding' : 'editing'} assignments. Ask a list owner to review your access.`, 'denied');
    }
    const api = this.listApi(this.config.authorizedAdminsListTitle);
    const [user, lookup, groupList] = await Promise.all([
      this.get<IAssignmentField>(`${api}/fields/getbyinternalnameortitle('User')`),
      this.get<IAssignmentField>(`${api}/fields/getbyinternalnameortitle('OfficeGroupRecord')`),
      this.get<{ Id: string }>(`${this.listApi(this.config.groupListTitle)}?$select=Id`)
    ]);
    const normalizeId = (value: string | undefined): string => (value || '').replace(/[{}]/g, '').toLowerCase();
    if (user.TypeAsString !== 'User' || user.ReadOnlyField || lookup.ReadOnlyField ||
        (lookup.TypeAsString !== 'LookupMulti' && lookup.TypeAsString !== 'Lookup') ||
        !lookup.LookupList || normalizeId(lookup.LookupList) !== normalizeId(groupList.Id) || !metadata.ListItemEntityTypeFullName) {
      throw new Error('This list’s field configuration does not support the assignment editor. Ask a list owner to review User and OfficeGroupRecord.');
    }
    return { metadata, user, lookup };
  }

  public async searchAdmins(query: string): Promise<IUser[]> {
    await this.requireAdminGroup();
    if (!this.canManage(await this.metadata())) throw new AssignmentSaveError('Your list permissions no longer allow access administration.', 'denied');
    const users = await new GraphService(this.context).searchUsers(query);
    return users.filter((user) => !user.isGroup && !!user.userPrincipalName);
  }

  public async openEditor(id?: number): Promise<IAssignmentEditor> {
    if (id !== undefined && (!Number.isSafeInteger(id) || id <= 0)) throw new Error('Invalid assignment record.');
    const schema = await this.editorSchema(id === undefined);
    const groups = await this.groups();
    const base = { groups, multiple: schema.lookup.TypeAsString === 'LookupMulti', required: schema.lookup.Required };
    if (id === undefined) return { ...base, userName: '', userEmail: '', groupIds: [] };
    const response = await this.context.spHttpClient.get(
      `${this.listApi(this.config.authorizedAdminsListTitle)}/items(${id})?$select=Id,User/Id,User/Title,User/EMail,OfficeGroupRecord/Id&$expand=User,OfficeGroupRecord`,
      SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=minimalmetadata', 'odata-version': '' } });
    if (!response.ok) throw new Error(`This assignment could not be opened (HTTP ${response.status}). Refresh assignments and check your permissions.`);
    const row: IAssignmentRow & { 'odata.etag'?: string; '@odata.etag'?: string } = await response.json();
    const etag = response.headers.get('ETag') || row['odata.etag'] || row['@odata.etag'];
    if (!etag || etag === '*') throw new Error('SharePoint did not return a record version. Refresh before editing.');
    if (!row.User || !row.User.Id) throw new Error('This record has no admin. Repair its User field in the assignments list before editing groups.');
    return { ...base, id, etag, userId: row.User.Id, userName: row.User.Title, userEmail: row.User.EMail || '', groupIds: lookupIds(row.OfficeGroupRecord) };
  }

  public async saveAssignment(editor: IAssignmentEditor, groupIds: number[], admin?: IUser): Promise<void> {
    const creating = editor.id === undefined;
    if (!creating && (!Number.isSafeInteger(editor.id) || (editor.id as number) <= 0 || !editor.etag || editor.etag === '*')) {
      throw new Error('Reopen the assignment to obtain its current version before saving.');
    }
    const schema = await this.editorSchema(creating);
    const ids = Array.from(new Set(groupIds));
    const groups = await this.groups();
    if (ids.some((id) => !Number.isSafeInteger(id) || id <= 0 || !groups.some((group) => group.id === id))) {
      throw new Error('One or more selected groups are no longer available. Remove unavailable groups or reload the assignment.');
    }
    const multiple = schema.lookup.TypeAsString === 'LookupMulti';
    if ((!multiple && ids.length > 1) || (schema.lookup.Required && !ids.length)) {
      throw new Error(multiple ? 'Select at least one group.' : 'This list requires a single group selection.');
    }
    const body: { [key: string]: unknown } = {
      __metadata: { type: schema.metadata.ListItemEntityTypeFullName },
      OfficeGroupRecordId: multiple ? { results: ids } : (ids[0] || null)
    };
    if (creating) {
      if (!admin || admin.isGroup || !admin.userPrincipalName) throw new Error('Select an individual admin from the directory results.');
      const response = await this.context.spHttpClient.post(`${this.webUrl}/_api/web/ensureuser`, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'odata-version': '' },
        body: JSON.stringify({ logonName: admin.userPrincipalName })
      });
      if (!response.ok) throw new Error(`SharePoint could not resolve this admin (HTTP ${response.status}). Check their account and your site access.`);
      const user: { Id: number; PrincipalType: number } = await response.json();
      if (!Number.isSafeInteger(user.Id) || user.Id <= 0 || user.PrincipalType !== 1) throw new Error('Select an individual user account.');
      // Honor a Person field restricted to a SharePoint group, without adding the user to that group.
      if (schema.user.SelectionGroup && schema.user.SelectionGroup > 0) {
        await this.get(`${this.webUrl}/_api/web/sitegroups(${schema.user.SelectionGroup})/users/getbyid(${user.Id})`);
      }
      body.UserId = user.Id;
      body.Title = admin.displayName;
    }
    const headers: { [key: string]: string } = {
      Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'odata-version': ''
    };
    if (!creating) { headers['X-HTTP-Method'] = 'MERGE'; headers['IF-MATCH'] = editor.etag as string; }
    let response;
    try {
      response = await this.context.spHttpClient.post(`${this.listApi(this.config.authorizedAdminsListTitle)}/items${creating ? '' : `(${editor.id})`}`,
        SPHttpClient.configurations.v1, { headers, body: JSON.stringify(body) });
    } catch {
      throw new AssignmentSaveError('The save could not be confirmed. Close the editor and refresh assignments before trying again; the change may already have saved.', 'uncertain');
    }
    if (response.ok) return;
    if (response.status === 412 || response.status === 409 || response.status === 404) {
      throw new AssignmentSaveError('This assignment changed or was removed while you were editing. Your draft is still here. Reload the assignment before saving again.', 'conflict');
    }
    if (response.status === 403 || response.status === 401) {
      throw new AssignmentSaveError('SharePoint denied this save. Your draft is still here. Ask a list owner to check your list and item permissions.', 'denied');
    }
    if (response.status >= 500) throw new AssignmentSaveError('SharePoint could not confirm the save. Close the editor and refresh assignments before trying again; the change may already have saved.', 'uncertain');
    throw new AssignmentSaveError(`SharePoint rejected the assignment (HTTP ${response.status}). Your draft is still here. Check the list’s required fields and validation rules.`, 'validation');
  }

  public async load(): Promise<IAccessOverview> {
    await this.requireAdminGroup();
    const metadata = await this.metadata();
    if (!this.canManage(metadata)) throw new Error('Managing access assignments requires Manage Lists and Edit Items on the Authorized Admins list.');
    const [rawAssignments, groups] = await Promise.all([
      this.all<IAssignmentRow>(`${this.listApi(this.config.authorizedAdminsListTitle)}/items?$select=Id,User/Id,User/Title,User/EMail,OfficeGroupRecord/Id&$expand=User,OfficeGroupRecord&$top=500`),
      this.groups()
    ]);
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
    return {
      assignments, groups, ...summarizeAssignments(assignments, groups),
      listUrl: this.formUrl(metadata.DefaultViewUrl), canAdd: this.hasPermission(metadata, SPPermission.addListItems)
    };
  }
}
