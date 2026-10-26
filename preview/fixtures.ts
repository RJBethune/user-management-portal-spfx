import { AccountManagementService } from '../src/webparts/accountManagement/services/AccountManagementService';
import { AccessAdministrationService } from '../src/webparts/accountManagement/services/AccessAdministrationService';
import { IAccessAssignment, IAssignmentEditor, summarizeAssignments } from '../src/webparts/accountManagement/models/accessAssignments';
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
  const searchPeople = async (query: string) => people.filter(person =>
    `${person.displayName} ${person.userPrincipalName}`.toLowerCase().includes(query.trim().toLowerCase()));
  const managedGroups = groups.map(group => ({ id: group.id, title: group.title, groupId: group.groupId, isTeamsConnected: false }));
  const assignments: IAccessAssignment[] = people.slice(0, 2).map((person, index) => ({
    id: index + 1, userId: Number(person.id), userName: person.displayName, userEmail: person.mail,
    groups: [managedGroups[index]], missingGroupIds: []
  }));
  // Only the local Portfolio runner installs these fixtures. Production service permissions remain intact.
  Object.assign(AccessAdministrationService.prototype, {
    isAvailable: async () => state.fixture !== 'not-admin',
    load: async () => {
      if (state.fixture === 'loading') return new Promise(() => {});
      if (state.fixture === 'error') throw Error('Fixture: access assignments could not be loaded.');
      const rows = state.fixture === 'empty' ? [] : assignments;
      const availableGroups = state.fixture === 'empty' ? [] : managedGroups;
      return { assignments: rows, groups: availableGroups, ...summarizeAssignments(rows, availableGroups), listUrl: '', canAdd: true };
    },
    openEditor: async (id?: number): Promise<IAssignmentEditor> => {
      const assignment = id === undefined ? undefined : assignments.find(row => row.id === id);
      if (id !== undefined && !assignment) throw Error('Fixture: assignment not found.');
      return {
        id: assignment?.id, etag: assignment ? '"1"' : undefined, userId: assignment?.userId,
        userName: assignment?.userName || '', userEmail: assignment?.userEmail || '',
        groupIds: assignment?.groups.map(group => group.id) || [],
        groups: state.fixture === 'empty' ? [] : managedGroups, multiple: true, required: false
      };
    },
    searchAdmins: searchPeople,
    saveAssignment: async () => { throw Error('Assignment writes require authenticated SharePoint preview. Local fixtures never contact live services.'); }
  });
  const service = AccountManagementService.prototype as any;
  Object.assign(service, {
    getAuthorizedGroups: async () => { if (state.fixture === 'loading') return new Promise(() => {}); if (state.fixture === 'error') throw Error('Fixture: authorized groups could not be loaded.'); return state.fixture === 'empty' ? [] : groups; },
    checkListHealth: async () => [],
    getRecentRequests: async () => {
      if (state.fixture === 'loading') return new Promise(() => {});
      if (state.fixture === 'error') throw Error('Fixture: recent requests could not be loaded.');
      return state.fixture === 'empty' ? [] : [
        { id: 52, action: 'Add Member', groupName: groups[0].title, groupId: groups[0].groupId, memberDisplayName: people[2].displayName, status: 'Completed', resultMessage: 'Member added.', requestedOn: '2026-10-26T12:00:00Z' },
        { id: 51, action: 'Remove Member', groupName: groups[1].title, groupId: groups[1].groupId, memberDisplayName: people[1].displayName, status: 'Failed', resultMessage: 'Workflow denied this change.', requestedOn: '2026-10-26T11:00:00Z' }
      ];
    }, getGroupSitePermissions: async () => [],
    createMembershipRequest: blocked, recordCompletedChange: blocked, changeSharePointGroupMembership: blocked, getRequest: blocked, pollRequest: blocked
  });
  Object.assign(GraphService.prototype, {
    getGroupMembers: async () => people, getGroupOwners: async () => people.slice(0, 1), getGroupPhotoUrl: async () => undefined,
    searchUsers: searchPeople, searchGroups: async () => []
  });
}
