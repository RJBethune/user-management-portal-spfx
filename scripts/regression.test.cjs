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
  static addListItems = 2;
  constructor(mask) { assert.equal(typeof mask.Low, 'number'); this.low = mask.Low; }
  hasPermission(permission) { return (this.low & permission) === permission; }
}
const { AccessAdministrationService } = load('src/webparts/accountManagement/services/AccessAdministrationService.ts', {
  '@microsoft/sp-http': { SPHttpClient: { configurations: { v1: {} } } },
  '@microsoft/sp-page-context': { SPPermission: Permission },
  './GraphService': { GraphService: class {} },
  '../models/accessAssignments': accessModels
});
const site = 'https://example.sharepoint.com/sites/access';
const config = { authorizedAdminsListTitle: "Admin's assignments", groupListTitle: 'Managed Groups' };
function administration(handler, post = async () => { throw new Error('Unexpected write'); }, listConfig = config) {
  const calls = [];
  const service = new AccessAdministrationService({
    pageContext: { web: { absoluteUrl: site } },
    spHttpClient: {
      get: async url => { calls.push(url); return { ok: true, headers: { get: () => '"3"' }, json: async () => handler(url) }; },
      post: async (url, _configuration, options) => { calls.push(url); return post(url, options); }
    }
  }, listConfig);
  return { service, calls };
}
function metadata(mask = '2054') {
  return { Id: 'list-id', ListItemEntityTypeFullName: 'SP.Data.AdminsListItem', EffectiveBasePermissions: { Low: mask, High: '0' },
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
test('assignment overview reads all pages, preserves duplicates, and targets the configured list', async () => {
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
  assert.equal(result.listUrl, `${site}/Lists/Admins/AllItems.aspx`);
  assert.equal(result.canAdd, true);
  assert.ok(calls.some(url => url.includes("Admin%27%27s") || url.includes("Admin''s")));
});

const groupRows = [{ Id: 1, Title: 'Operations', GroupId: 'group-1' }, { Id: 2, Title: 'Coordination', GroupId: 'group-2' }];
const editor = { id: 7, etag: '"3"', userId: 10, userName: 'Admin', userEmail: 'admin@example.test', groupIds: [1], groups: [], multiple: true, required: false };
const newAdmin = { id: 'entra-guid', displayName: 'New admin', userPrincipalName: 'new@example.test' };
function editorReads({ mask = '2054', type = 'LookupMulti', required = false, lookupList = '{groups-id}', selectionGroup = 0, readOnly = false } = {}) {
  return url => {
    if (url.includes('EffectiveBasePermissions')) return metadata(mask);
    if (url.includes("getbyinternalnameortitle('User')")) return { TypeAsString: 'User', ReadOnlyField: false, SelectionGroup: selectionGroup };
    if (url.includes("getbyinternalnameortitle('OfficeGroupRecord')")) return { TypeAsString: type, ReadOnlyField: readOnly, Required: required, LookupList: lookupList };
    if (url.endsWith('?$select=Id')) return { Id: 'groups-id' };
    if (url.includes('/items(7)')) return { Id: 7, User: { Id: 10, Title: 'Admin' }, OfficeGroupRecord: [{ Id: 1 }] };
    if (url.includes('/sitegroups')) throw new Error('The selected admin is outside the Person field’s allowed group.');
    return { value: groupRows };
  };
}
test('editor reads the current item version and saves only its groups with conditional MERGE', async () => {
  let writes = 0;
  const { service } = administration(editorReads(), async (url, options) => {
    writes++;
    assert.ok(url.endsWith('/items(7)'));
    assert.equal(options.headers['IF-MATCH'], '"3"');
    assert.equal(options.headers['X-HTTP-Method'], 'MERGE');
    assert.equal(options.headers['odata-version'], '');
    assert.deepEqual(JSON.parse(options.body), { __metadata: { type: 'SP.Data.AdminsListItem' }, OfficeGroupRecordId: { results: [2] } });
    return { ok: true, status: 204 };
  });
  const opened = await service.openEditor(7);
  assert.equal(opened.etag, '"3"'); assert.deepEqual(opened.groupIds, [1]);
  await service.saveAssignment(opened, [2, 2]);
  assert.equal(writes, 1);
});
test('every save rechecks list management permission, including Add Items for creation', async () => {
  for (const mask of ['0', '4', '2048']) {
    const { service, calls } = administration(editorReads({ mask }));
    await assert.rejects(service.saveAssignment(editor, [1]), /permissions no longer/);
    assert.equal(calls.length, 1);
  }
  const { service } = administration(editorReads({ mask: '2052' }));
  await assert.rejects(service.openEditor(), /permissions no longer allow adding/);
  await assert.rejects(service.saveAssignment({ ...editor, id: undefined }, [1], newAdmin), /permissions no longer allow adding/);
  assert.equal((await service.load()).canAdd, false);
});
test('new assignments resolve the site-local user ID and preserve the configured list site', async () => {
  const configuredSite = 'https://example.sharepoint.com/sites/configured';
  const writes = [];
  const { service } = administration(editorReads(), async (url, options) => {
    assert.ok(url.startsWith(configuredSite + '/_api/'));
    writes.push(JSON.parse(options.body));
    if (url.endsWith('/ensureuser')) return { ok: true, json: async () => ({ Id: 42, PrincipalType: 1 }) };
    assert.equal(options.headers['IF-MATCH'], undefined);
    return { ok: true, status: 201 };
  }, { ...config, listSiteUrl: configuredSite });
  await service.saveAssignment({ ...editor, id: undefined }, [1, 2], newAdmin);
  assert.deepEqual(writes, [{ logonName: newAdmin.userPrincipalName }, {
    __metadata: { type: 'SP.Data.AdminsListItem' }, OfficeGroupRecordId: { results: [1, 2] }, UserId: 42, Title: 'New admin'
  }]);
});
test('single-value lookups use scalar IDs; clearing optional multi-value lookups uses an empty collection', async () => {
  for (const [type, ids, expected] of [['Lookup', [2], 2], ['Lookup', [], null], ['LookupMulti', [], { results: [] }]]) {
    const { service } = administration(editorReads({ type }), async (_url, options) => {
      assert.deepEqual(JSON.parse(options.body).OfficeGroupRecordId, expected); return { ok: true };
    });
    await service.saveAssignment(editor, ids);
  }
});
test('invalid IDs, required groups, incompatible schemas and missing versions never reach a write', async () => {
  const cases = [
    [{}, [99]], [{}, [-1]], [{}, ['1']], [{ required: true }, []], [{ type: 'Lookup' }, [1, 2]],
    [{ type: 'Text' }, [1]], [{ lookupList: 'unrelated-list' }, [1]], [{ readOnly: true }, [1]]
  ];
  for (const [settings, ids] of cases) {
    const { service } = administration(editorReads(settings));
    await assert.rejects(service.saveAssignment(editor, ids));
  }
  const { service } = administration(editorReads());
  for (const etag of ['', undefined, '*']) await assert.rejects(service.saveAssignment({ ...editor, etag }, [1]), /current version/);
});
test('SharePoint item denial and conflicts preserve errors without retries or wildcard overwrites', async () => {
  for (const [status, code] of [[403, 'denied'], [401, 'denied'], [409, 'conflict'], [412, 'conflict'], [404, 'conflict'], [400, 'validation'], [500, 'uncertain']]) {
    let writes = 0;
    const { service } = administration(editorReads(), async () => { writes++; return { ok: false, status }; });
    await assert.rejects(service.saveAssignment(editor, [2]), error => error.code === code);
    assert.equal(writes, 1);
  }
  const { service } = administration(editorReads(), async () => { throw new Error('Connection lost'); });
  await assert.rejects(service.saveAssignment(editor, [2]), error => error.code === 'uncertain');
});
test('new admins honor Person-field group restrictions and reject group principals', async () => {
  for (const [selectionGroup, principalType] of [[12, 1], [0, 4]]) {
    let writes = 0;
    const { service } = administration(editorReads({ selectionGroup }), async (url) => {
      writes++; assert.ok(url.endsWith('/ensureuser'));
      return { ok: true, json: async () => ({ Id: 42, PrincipalType: principalType }) };
    });
    await assert.rejects(service.saveAssignment({ ...editor, id: undefined }, [1], newAdmin));
    assert.equal(writes, 1);
  }
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
