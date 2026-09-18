'use strict';
const fs=require('node:fs'),path=require('node:path');
const root=path.resolve(__dirname,'..');
const read=relative=>JSON.parse(fs.readFileSync(path.join(root,relative),'utf8'));
const app=read('delivery/app.json');
const tool=require('@eo/spfx-baseline/contracts/toolchain.json').lanes.heft;
const ui=require('@eo/spfx-baseline/contracts/ui-runtime.json');
const host=read((app.projectRoot==='.'?'':app.projectRoot+'/')+'package.json');
for(const [name,expected] of Object.entries({'@rushstack/heft':tool.heft,'@microsoft/sp-core-library':tool.spfx,react:tool.react,typescript:tool.typescript}))if((host.dependencies?.[name]||host.devDependencies?.[name])!==expected)throw Error(name+' differs from the pinned baseline');
const pkg=read('package.json'),lock=read('package-lock.json');
for(const [name,version] of Object.entries(ui.directDependencies))if(pkg.dependencies[name]!==version)throw Error('Direct pin differs: '+name);
for(const [name,version] of Object.entries(ui.overrides)){
 if(pkg.overrides[name]!==version)throw Error('Override differs: '+name);
 for(const [relative,record] of Object.entries(lock.packages))if(relative.endsWith('node_modules/'+name)){
  if(record.version!==version||read(relative+'/package.json').version!==version)throw Error('Installed/locked closure differs: '+name);
 }
}
for(const name of ui.singlePhysicalPackages){const entries=Object.keys(lock.packages).filter(n=>n.endsWith('node_modules/'+name));if(entries.length!==1||entries[0]!=='node_modules/'+name)throw Error('Expected one top-level copy: '+name);}
console.log('Pinned toolchain and complete installed Fluent/Tabster closure verified.');
