const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
function load(relative) {
  const file = path.join(__dirname, '..', relative);
  const source = ts.transpileModule(fs.readFileSync(file, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText;
  const module = { exports: {} };
  new Function('module', 'exports', 'require', source)(module, module.exports, require);
  return module.exports;
}
const { isSharePointGroup } = load('src/webparts/accountManagement/shared/groupType.ts');
const { getManageability } = load('src/webparts/accountManagement/shared/manageability.ts');
test('group routing separates SharePoint integer IDs from Microsoft 365 GUIDs', () => {
  assert.equal(isSharePointGroup(' 42 '), true);
  assert.equal(isSharePointGroup('762df7fa-0fd0-4019-b227-d184ab4abc67'), false);
  assert.equal(isSharePointGroup(undefined), false);
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
