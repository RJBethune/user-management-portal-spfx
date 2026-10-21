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
const members = [
  { id: 'person-1', displayName: 'Taylor Chen', mail: 'taylor.chen@example.test', jobTitle: 'Operations coordinator' },
  { id: 'person-2', displayName: 'Jordan Patel', mail: 'jordan.patel@example.test', jobTitle: 'Program officer' }
];
const context = {
  pageContext: { web: { absoluteUrl: site }, site: { id: 'preview-site' }, user: { loginName: 'alex@example.test', displayName: 'Alex Morgan' } },
  spHttpClient: {
    get: async (raw: string) => {
      const url = decodeURIComponent(raw);
      if (scenario === 'error' && url.includes('User/Title')) return { ok: false, status: 403, json: async () => ({}) };
      let result: unknown = { value: [] };
      if (url.includes('EffectiveBasePermissions')) result = {
        Id: 'admins', EffectiveBasePermissions: { Low: scenario === 'denied' ? '4' : '2052', High: '0' },
        DefaultViewUrl: '/sites/access/Lists/Admins/AllItems.aspx', DefaultEditFormUrl: '/sites/access/Lists/Admins/EditForm.aspx', DefaultNewFormUrl: '/sites/access/Lists/Admins/NewForm.aspx'
      };
      else if (url.includes('/fields')) {
        const titles = [DEFAULT_LIST_CONFIG.groupListTitle, DEFAULT_LIST_CONFIG.requestListTitle, DEFAULT_LIST_CONFIG.authorizedAdminsListTitle, DEFAULT_LIST_CONFIG.sitePermissionsListTitle];
        const spec = EXPECTED_LISTS.find(item => url.includes("getbytitle('" + ({ groups: titles[0], requests: titles[1], admins: titles[2], sitePermissions: titles[3] } as Record<string,string>)[item.key] + "')"));
        result = { value: spec?.columns.map(column => ({ InternalName: column.name, TypeAsString: column.types[0] })) || [] };
      }
      else if (url.includes('/currentuser')) result = { Id: 10, Title: 'Alex Morgan', Email: 'alex@example.test' };
      else if (url.includes("getbytitle('Managed Groups')")) result = { value: url.includes('$filter=') ? groups.slice(0, 2) : groups };
      else if (url.includes("getbytitle('Group Management Authorized Admins')")) result = { value: url.includes('UserId eq') ? (scenario === 'empty' ? [] : assignments.slice(0, 1)) : assignments };
      return { ok: true, status: 200, json: async () => result, text: async () => JSON.stringify(result) };
    },
    post: async () => { throw new Error('Preview writes are disabled.'); }
  },
  msGraphClientFactory: { getClient: async () => ({ api: (path: string) => {
    const request = { header: () => request, responseType: () => request, get: async () => ({ value: path.includes('/owners') ? members.slice(0, 1) : members }) };
    return request;
  } }) }
};
function mount(instanceId: string, mountPoint: HTMLElement) {
  const host = createSpfxUiHost({ mountPoint, portalParent: document.body, targetDocument: document, instanceId, style: 'luma', theme: initialTokens });
  host.portalHost.style.position = 'relative'; host.portalHost.style.zIndex = '1000';
  ReactDOM.render(<SpfxUiHostProvider host={host}><AccountManagementShell host={host} context={context as any} description="365 Account Management" userDisplayName="Alex Morgan" buildVersion="1.12.0" listConfig={DEFAULT_LIST_CONFIG} visibleOffices="" pollTimeoutMs={120000} introText="" helpText="Contact the service desk for membership support." helpUrl="" startCollapsed={false} requireJustification={true} showGroupPhotos={false} isEditMode={false} /></SpfxUiHostProvider>, host.appRoot);
  return () => { ReactDOM.unmountComponentAtNode(host.appRoot); host.dispose(); };
}
const dispose = mount('preview-account-management', document.getElementById('app')!);
(window as any).disposePreview = dispose;
if (new URLSearchParams(location.search).has('adjacent')) {
  const adjacent = document.createElement('div'); document.body.appendChild(adjacent); mount('preview-adjacent', adjacent);
}
