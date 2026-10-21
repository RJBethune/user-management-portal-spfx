const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
function load(relative, imports = {}) {
  const file = path.join(__dirname, '..', relative);
  const source = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText;
  const module = { exports: {} };
  new Function('module', 'exports', 'require', source)(module, module.exports, name => imports[name] || require(name));
  return module.exports;
}
const { isSharePointGroup } = load('src/webparts/accountManagement/shared/groupType.ts');
const { getManageability } = load('src/webparts/accountManagement/shared/manageability.ts');
test('group routing separates SharePoint integer IDs from Microsoft 365 GUIDs', () => {
  assert.equal(isSharePointGroup(' 42 '), true);
  assert.equal(isSharePointGroup('762df7fa-0fd0-4019-b227-d184ab4abc67'), false);
  assert.equal(isSharePointGroup(undefined), false);
});

const accessModels = load('src/webparts/accountManagement/models/accessAssignments.ts');
test('assignment lookup accepts historical and multi-value records without duplicate or invalid IDs', () => {
  for (const value of [{ Id: 12 }, [{ Id: 12 }], { results: [{ Id: 12 }, { Id: 12 }, { Id: -1 }, { Id: '13' }] }]) {
    assert.deepEqual(accessModels.lookupIds(value), [12]);
  }
  assert.deepEqual(accessModels.lookupIds(null), []);
});

test('coverage excludes missing admins and flags duplicate assignment rows', () => {
  const groups = [{ id: 1 }, { id: 2 }, { id: 3 }];
  const result = accessModels.summarizeAssignments([
    { userId: 4, groups: [groups[0]] }, { userId: 4, groups: [] }, { groups: [groups[1]] }
  ], groups);
  assert.deepEqual(result.unassigned, [groups[1], groups[2]]);
  assert.deepEqual(result.duplicateUserIds, [4]);
});

// This boundary double also verifies conversion of SharePoint's string-valued permission mask.
class Permission {
  static manageLists = 2048;
  static editListItems = 4;
  constructor(mask) { assert.equal(typeof mask.Low, 'number'); this.low = mask.Low; }
  hasPermission(permission) { return (this.low & permission) === permission; }
}
const { AccessAdministrationService } = load('src/webparts/accountManagement/services/AccessAdministrationService.ts', {
  '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
  '@microsoft/sp-page-context': { SPPermission: Permission },
  '../models/accessAssignments': accessModels
});
const site = 'https://example.sharepoint.com/sites/access';
const config = { authorizedAdminsListTitle: "Admin's assignments", groupListTitle: 'Managed Groups' };
function administration(handler) {
  const calls = [];
  const service = new AccessAdministrationService({
    pageContext: { web: { absoluteUrl: site } },
    spHttpClient: { get: async url => { calls.push(url); return { ok: true, json: async () => handler(url) }; } }
  }, config);
  return { service, calls };
}
function metadata(mask = '2052') {
  return { Id: 'list-id', EffectiveBasePermissions: { Low: mask, High: '0' },
    DefaultViewUrl: '/sites/access/Lists/Admins/AllItems.aspx',
    DefaultEditFormUrl: '/sites/access/Lists/Admins/EditForm.aspx',
    DefaultNewFormUrl: '/sites/access/Lists/Admins/NewForm.aspx' };
}
test('assignment overview fails closed before enumerating records without management permission', async () => {
  for (const mask of ['0', '4', '2048']) {
    const { service, calls } = administration(() => metadata(mask));
    assert.equal(await service.isAvailable(), false);
    await assert.rejects(service.load(), /Manage Lists and Edit Items/);
    assert.ok(calls.every(url => !url.includes('/items?')));
  }
});
test('assignment overview reads all pages, preserves duplicates, and targets actual list forms', async () => {
  const { service, calls } = administration(url => {
    if (url.includes('EffectiveBasePermissions')) return metadata();
    if (url.includes('Managed%20Groups')) return { value: [{ Id: 1, Title: 'Operations', GroupId: 'group-1' }, { Id: 2, Title: 'Unassigned', GroupId: 'group-2' }] };
    if (url.includes('page=2')) return { value: [{ Id: 8, User: { Id: 10, Title: 'Admin' }, OfficeGroupRecord: { Id: 99 } }] };
    return { value: [{ Id: 7, User: { Id: 10, Title: 'Admin' }, OfficeGroupRecord: [{ Id: 1 }] }], 'odata.nextLink': `${site}/_api/assignments?page=2` };
  });
  const result = await service.load();
  assert.equal(result.assignments.length, 2);
  assert.deepEqual(result.duplicateUserIds, [10]);
  assert.deepEqual(result.assignments[1].missingGroupIds, [99]);
  assert.equal(result.unassigned[0].id, 2);
  assert.equal(result.editItemUrl(7), `${site}/Lists/Admins/EditForm.aspx?ID=7`);
  assert.ok(calls.some(url => url.includes("Admin%27%27s") || url.includes("Admin''s")));
});
test('assignment overview refuses foreign continuations instead of returning partial results', async () => {
  const { service } = administration(url => url.includes('EffectiveBasePermissions') ? metadata() : {
    value: [], '@odata.nextLink': 'https://other.example/_api/items'
  });
  await assert.rejects(service.load(), /Invalid assignment list continuation/);
});
test('malformed and restricted groups cannot enable membership actions', () => {
  for (const groupId of ['', 'not-a-group', '42x']) assert.equal(getManageability({ groupId }).manageable, false);
  const groupId = '762df7fa-0fd0-4019-b227-d184ab4abc67';
  for (const groupType of ['Dynamic', 'Mail Enabled Security', 'Distribution', 'RoleAssignable', 'On Premises']) {
    assert.equal(getManageability({ groupId, groupType }).manageable, false, groupType);
  }
  assert.equal(getManageability({ groupId, groupType: 'Microsoft365' }).manageable, true);
  assert.equal(getManageability({ groupId: '42' }).manageable, true);
});
