import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { createSpfxUiHost, SpfxUiHostProvider } from '@spfx-kit/ui-profile';
import { initialTokens } from '../src/ui/theme';
import '../src/ui/styles';
import { AccountManagementShell } from '../src/webparts/accountManagement/components/AccountManagementShell';
import { DEFAULT_LIST_CONFIG } from '../src/webparts/accountManagement/services/AccountManagementService';
import { EXPECTED_LISTS } from '../src/webparts/accountManagement/models/listSchema';

const scenario = new URLSearchParams(location.search).get('scenario');
const site = 'https://example.sharepoint.com/sites/access';
const groups = [
  { Id: 1, Title: 'Operations team', GroupId: '12345678-1234-1234-1234-123456789012', IsTeamsConnected: true, Description: 'Membership for the Operations team.', Mail: 'operations@example.test', Visibility: 'Private' },
  { Id: 2, Title: 'Regional coordination and emergency planning', GroupId: '22345678-1234-1234-1234-123456789012', IsTeamsConnected: false },
  { Id: 3, Title: 'New office', GroupId: '32345678-1234-1234-1234-123456789012', IsTeamsConnected: false }
];
const assignments = [
  { Id: 7, User: { Id: 10, Title: 'Alex Morgan', EMail: 'alex.morgan@example.test' }, OfficeGroupRecord: [{ Id: 1 }, { Id: 2 }] },
  { Id: 8, User: { Id: 11, Title: 'Sam Rivera', EMail: 'sam.rivera@example.test' }, OfficeGroupRecord: [{ Id: 1 }] },
  { Id: 9, User: { Id: 10, Title: 'Alex Morgan', EMail: 'alex.morgan@example.test' }, OfficeGroupRecord: [] }
];
if (scenario === 'workspace') {
  ['Finance and budget', 'Employee onboarding', 'IT service coordination', 'Leadership team'].forEach((title, index) => {
    const Id = index + 4;
    groups.push({ Id, Title: title, GroupId: `${Id}2345678-1234-1234-1234-123456789012`, IsTeamsConnected: index !== 2 });
    assignments[0].OfficeGroupRecord.push({ Id });
  });
}
const members = [
  { id: 'person-1', displayName: 'Taylor Chen', userPrincipalName: 'taylor.chen@example.test', mail: 'taylor.chen@example.test', jobTitle: 'Operations coordinator' },
  { id: 'person-2', displayName: 'Jordan Patel', userPrincipalName: 'jordan.patel@example.test', mail: 'jordan.patel@example.test', jobTitle: 'Program officer' }
];
const versions = new Map<number, number>();
const writes: unknown[] = [];
(window as any).previewWrites = writes;
let conflictShown = false;
let resolvedAdmin = members[0];
(window as any).previewAdminMember = scenario !== 'not-admin';
const context = {
  pageContext: { web: { absoluteUrl: site }, site: { id: 'preview-site' }, user: { loginName: 'alex@example.test', displayName: 'Alex Morgan' } },
  spHttpClient: {
    get: async (raw: string) => {
      const url = decodeURIComponent(raw);
      if (scenario === 'error' && url.includes('User/Title')) return { ok: false, status: 403, json: async () => ({}) };
      let result: unknown = { value: [] };
      if (url.includes('/currentuser/groups')) result = { value: (window as any).previewAdminMember ? [{ Id: 42, Title: '365AccessManagement-Admins' }] : [] };
      else if (url.includes('EffectiveBasePermissions')) result = {
        Id: 'admins', ListItemEntityTypeFullName: 'SP.Data.AdminsListItem', EffectiveBasePermissions: { Low: scenario === 'denied' ? '4' : scenario === 'noadd' || (scenario === 'inventory-change' && writes.length > 0) ? '2052' : '2054', High: '0' },
        DefaultViewUrl: '/sites/access/Lists/Admins/AllItems.aspx', DefaultEditFormUrl: '/sites/access/Lists/Admins/EditForm.aspx', DefaultNewFormUrl: '/sites/access/Lists/Admins/NewForm.aspx'
      };
      else if (url.includes("/fields/getbyinternalnameortitle('User')")) result = { TypeAsString: 'User', ReadOnlyField: false, SelectionGroup: 0 };
      else if (url.includes("/fields/getbyinternalnameortitle('OfficeGroupRecord')")) result = { TypeAsString: scenario === 'single' ? 'Lookup' : 'LookupMulti', ReadOnlyField: false, Required: false, LookupList: '{managed-groups}' };
      else if (url.includes('/fields')) {
        const titles = [DEFAULT_LIST_CONFIG.groupListTitle, DEFAULT_LIST_CONFIG.requestListTitle, DEFAULT_LIST_CONFIG.authorizedAdminsListTitle, DEFAULT_LIST_CONFIG.sitePermissionsListTitle];
        const spec = EXPECTED_LISTS.find(item => url.includes("getbytitle('" + ({ groups: titles[0], requests: titles[1], admins: titles[2], sitePermissions: titles[3] } as Record<string,string>)[item.key] + "')"));
        result = { value: spec?.columns.map(column => ({ InternalName: column.name, TypeAsString: column.types[0] })) || [] };
      }
      else if (url.includes('/currentuser')) result = { Id: 10, Title: 'Alex Morgan', Email: 'alex@example.test' };
      else if (url.includes("getbytitle('Group Membership Requests')/items")) {
        if (scenario === 'requests-error') return { ok: false, status: 403, json: async () => ({}) };
        result = { value: scenario === 'requests' ? [
          { Id: 52, Action: 'Add Member', GroupName: 'Operations team', GroupId: groups[0].GroupId, MemberDisplayName: 'Taylor Chen', Status: 'Completed', ResultMessage: 'Member added.', RequestedOn: '2026-10-26T12:00:00Z' },
          { Id: 51, Action: 'Remove Member', GroupName: 'Regional coordination', GroupId: groups[1].GroupId, MemberDisplayName: 'Jordan Patel', Status: 'Failed', ResultMessage: 'Workflow denied this change.' },
          { Id: 50, Action: 'Add Member', GroupName: 'Operations team', GroupId: groups[0].GroupId, MemberDisplayName: 'Alex Morgan', Status: 'Pending', ResultMessage: 'Waiting for the membership workflow.', RequestedOn: new Date().toISOString() }
        ] : [] };
      }
      else if (url.includes("getbytitle('Managed Groups')")) result = url.endsWith('?$select=Id') ? { Id: 'managed-groups' } : { value: url.includes('$filter=') ? groups.filter(row => assignments[0].OfficeGroupRecord.some(group => group.Id === row.Id)) : groups };
      else if (url.includes("getbytitle('Group Management Authorized Admins')")) {
        const match = url.match(/\/items\((\d+)\)/);
        result = match ? assignments.find(row => row.Id === Number(match[1])) : { value: url.includes('UserId eq') ? (scenario === 'empty' ? [] : assignments.slice(0, 1)) : assignments };
      }
      const id = Number(url.match(/\/items\((\d+)\)/)?.[1]);
      return { ok: true, status: 200, headers: { get: () => `"${versions.get(id) || 1}"` }, json: async () => result, text: async () => JSON.stringify(result) };
    },
    post: async (raw: string, _configuration: unknown, options: { headers: Record<string, string>; body: string }) => {
      const url = decodeURIComponent(raw);
      const body = JSON.parse(options.body);
      if (url.endsWith('/ensureuser')) {
        resolvedAdmin = members.find(person => person.userPrincipalName === body.logonName)!;
        return { ok: !!resolvedAdmin, status: resolvedAdmin ? 200 : 400, json: async () => ({ Id: 20 + members.indexOf(resolvedAdmin), PrincipalType: 1 }) };
      }
      if (!url.includes("getbytitle('Group Management Authorized Admins')/items")) throw new Error('Tenant writes are disabled in this preview.');
      writes.push({ url, body, headers: options.headers });
      if (scenario === 'pending-save') await new Promise<void>((resolve) => { (window as any).finishPreviewSave = resolve; });
      if (scenario === 'write-denied') return { ok: false, status: 403 };
      if (scenario === 'uncertain') throw new Error('Connection lost');
      const id = Number(url.match(/\/items\((\d+)\)/)?.[1]);
      if (scenario === 'conflict' && !conflictShown) {
        conflictShown = true;
        const row = assignments.find(item => item.Id === id)!;
        row.OfficeGroupRecord = [{ Id: 3 }]; versions.set(id, 2);
        return { ok: false, status: 412 };
      }
      if (id && options.headers['IF-MATCH'] !== `"${versions.get(id) || 1}"`) return { ok: false, status: 412 };
      const selected = Array.isArray(body.OfficeGroupRecordId?.results) ? body.OfficeGroupRecordId.results : body.OfficeGroupRecordId ? [body.OfficeGroupRecordId] : [];
      if (id) { assignments.find(row => row.Id === id)!.OfficeGroupRecord = selected.map((groupId: number) => ({ Id: groupId })); versions.set(id, (versions.get(id) || 1) + 1); }
      else assignments.push({ Id: 10 + assignments.length, User: { Id: body.UserId, Title: resolvedAdmin.displayName, EMail: resolvedAdmin.mail }, OfficeGroupRecord: selected.map((groupId: number) => ({ Id: groupId })) });
      if (scenario === 'inventory-change') assignments.splice(assignments.findIndex(row => row.Id === 8), 1);
      return { ok: true, status: id ? 204 : 201 };
    }
  },
  msGraphClientFactory: { getClient: async () => ({ api: (path: string) => {
    const request = { header: () => request, responseType: () => request, get: async () => ({ value: path.includes('/owners') ? members.slice(0, 1) : members }) };
    return request;
  } }) }
};
function mount(instanceId: string, mountPoint: HTMLElement) {
  const host = createSpfxUiHost({ mountPoint, portalParent: document.body, targetDocument: document, instanceId, style: 'luma', theme: initialTokens });
  host.portalHost.style.position = 'relative'; host.portalHost.style.zIndex = '1000';
  ReactDOM.render(<SpfxUiHostProvider host={host}><AccountManagementShell host={host} context={context as any} description="365 Account Management" userDisplayName="Alex Morgan" buildVersion="1.16.0" listConfig={DEFAULT_LIST_CONFIG} visibleOffices="" pollTimeoutMs={120000} introText="" helpText="Contact the service desk for membership support." helpUrl="" startCollapsed={scenario === 'start-collapsed'} requireJustification={true} showGroupPhotos={false} isEditMode={false} /></SpfxUiHostProvider>, host.appRoot);
  return () => { ReactDOM.unmountComponentAtNode(host.appRoot); host.dispose(); };
}
const dispose = mount('preview-account-management', document.getElementById('app')!);
(window as any).disposePreview = dispose;
if (new URLSearchParams(location.search).has('adjacent')) {
  const adjacent = document.createElement('div'); document.body.appendChild(adjacent); mount('preview-adjacent', adjacent);
}
