import { AccountManagementService } from '../src/webparts/accountManagement/services/AccountManagementService';
import { GraphService } from '../src/webparts/accountManagement/services/GraphService';
export function configure(properties: any) { /* Keep the app's actual settings defaults. */ }
export function install(state: any) {
  const blocked = async () => { throw Error('Membership writes require authenticated SharePoint preview. Local fixtures never contact live services.'); };
  const groups = [
    { id: 1, title: 'Executive Office', groupId: '762df7fa-0fd0-4019-b227-d184ab4abc67', groupType: 'Microsoft365', description: 'Sample Microsoft 365 group', siteUrl: '' },
    { id: 2, title: 'Site Contributors', groupId: '42', groupType: 'SharePoint', description: 'Sample site group', siteUrl: 'https://example.test/sites/preview' },
    { id: 3, title: 'Dynamic Office Group', groupId: '8a9d9e4f-a21d-4006-8615-bd2b64a9dcd9', groupType: 'Dynamic', description: 'Rule-based membership example', siteUrl: '' }
  ];
  const people = ['Alex Morgan', 'Jordan Lee', 'Taylor Davis'].map((displayName, i) => ({ id: String(i + 1), displayName, mail: 'person' + i + '@example.test', userPrincipalName: 'person' + i + '@example.test' }));
  const service = AccountManagementService.prototype as any;
  Object.assign(service, {
    getAuthorizedGroups: async () => { if (state.fixture === 'loading') return new Promise(() => {}); if (state.fixture === 'error') throw Error('Fixture: authorized groups could not be loaded.'); return state.fixture === 'empty' ? [] : groups; },
    checkListHealth: async () => [], getRecentRequests: async () => [], getGroupSitePermissions: async () => [],
    createMembershipRequest: blocked, recordCompletedChange: blocked, changeSharePointGroupMembership: blocked, getRequest: blocked, pollRequest: blocked
  });
  Object.assign(GraphService.prototype, {
    getGroupMembers: async () => people, getGroupOwners: async () => people.slice(0, 1), getGroupPhotoUrl: async () => undefined,
    searchUsers: async (query: string) => people.filter(p => p.displayName.toLowerCase().includes(query.toLowerCase())), searchGroups: async () => []
  });
}
