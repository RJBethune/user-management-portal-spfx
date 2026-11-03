import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import crypto from 'node:crypto';

// The bundled Codex presentation runtime supplies Artifact Tool. No app dependency changes.
const source = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(source, '../../..');
const scratch = path.join(root, 'scratch/rollout-build');
const modules = process.env.ARTIFACT_NODE_MODULES;
const skill = process.env.PRESENTATIONS_SKILL;
const python = process.env.ARTIFACT_PYTHON;
if (!modules || !skill || !python) throw new Error('Set ARTIFACT_NODE_MODULES, PRESENTATIONS_SKILL, and ARTIFACT_PYTHON. See docs/rollout/README.md.');
process.env.RUNTIME_NODE_MODULES = modules;
const require = createRequire(import.meta.url);
const { FileBlob, PresentationFile } = await import(pathToFileURL(require.resolve('@oai/artifact-tool', { paths: [modules] })).href);
const { finalizePresentation } = await import(pathToFileURL(path.join(skill, 'container_tools/artifact_tool_utils.mjs')).href);
const data = JSON.parse(await fs.readFile(path.join(source, 'plan.json'), 'utf8'));
const template = path.join(source, 'brand-template.pptx');
const p = await PresentationFile.importPptx(await FileBlob.load(template));
p.slides.keep([]);
const C = { navy:'#0A2240', blue:'#002D74', teal:'#0A314D', red:'#981B1E', crimson:'#D01319', ink:'#1B2027', muted:'#56606C', panel:'#F2F5F8', rule:'#D4DAE2', white:'#FFFFFF' };
const M=57.6, W=1164.8;
const tableSlides=[];
await fs.mkdir(path.join(scratch,'rendered'),{recursive:true});

function rect(s,name,x,y,w,h,fill,r=0,line='none') {
  return s.shapes.add({name,geometry:r?'roundRect':'rect',position:{left:x,top:y,width:w,height:h},fill,borderRadius:r,line:{fill:line,width:line==='none'?0:1}});
}
function text(s,name,value,x,y,w,h,size=23,opts={}) {
  const sh=s.shapes.add({name,geometry:'textbox',position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
  sh.text.style={fontSize:size,typeface:opts.head?'Garamond':'Calibri',color:opts.color||C.ink,bold:!!opts.bold,insets:0,wrap:'square',autoFit:'none',verticalAlignment:'top',alignment:opts.align||'left'};
  sh.text=value;
  return sh;
}
function ph(s,type,value,size,color) {
  const item=s.placeholders.items.find(x=>x.placeholderType===type);
  if(!item) throw new Error('Missing template placeholder '+type);
  item.text=value;
  item.text.style={fontSize:size,typeface:type==='title'?'Garamond':'Calibri',color,bold:true,insets:0,autoFit:'none',wrap:'square',verticalAlignment:'top'};
  return item;
}
function slide(d,layout='BA Content') {
  const s=p.slides.add({layout});
  s.placeholders.addFromLayout();
  ph(s,'title',d.title,layout==='BA Closing'?48:45.33,layout==='BA Closing'?C.white:C.navy);
  ph(s,'body',d.section,14.67,layout==='BA Closing'?C.white:C.navy);
  s.speakerNotes.text=d.notes;
  return s;
}
function lede(s,value){if(value)text(s,'Introduction',value,M,145,W,64,23);}
function paired(s,items,x,y,w,rowH,{size=22,headSize=23,color=C.blue}={}) {
  items.forEach(([heading,body],i)=>{
    const top=y+i*rowH;
    // A shared text box lets wrapping headings push the body down naturally.
    text(s,heading,[{runs:[{run:heading,style:{bold:true,color}}],spaceAfter:5},{runs:[body]}],x,top,w,rowH-10,size);
  });
}
async function icon(s,name,x,y) {
  rect(s,name+' disc',x,y,62,62,C.navy,31);
  s.images.add({blob:new Uint8Array(await fs.readFile(path.join(source,'images',name+'.png'))),contentType:'image/png',alt:'',position:{left:x+14,top:y+14,width:34,height:34}});
}

// Cover: keep the supplied layout's flag, department masthead, navy field, and red rule.
{
  const s=p.slides.add({layout:'BA Title'});s.placeholders.addFromLayout();
  const title=ph(s,'title',data.title,65,C.white);
  title.position={left:72,top:245,width:1136,height:95};
  const subtitle=ph(s,'body',data.subtitle,37.33,C.white);
  subtitle.position={left:72,top:350,width:1136,height:62};
  subtitle.text.style={typeface:'Garamond',fontSize:37.33,bold:false,color:C.white,insets:0};
  text(s,'Eyebrow','ROLLOUT GUIDE AND COMMUNICATION PLAN',72,194,1100,28,16,{bold:true,color:C.white});
  text(s,'Purpose','Who needs to know, what they should do, and how the office supports adoption.',72,433,1090,68,25,{color:C.white});
  text(s,'Audience','M EXECUTIVE OFFICE (M/EX)',72,548,1080,35,18,{bold:true,color:C.white});
  text(s,'Version','Account Management '+data.version+'\nProposed plan. Communications determines the schedule.',72,651,720,53,14.67);
  text(s,'Ownership','Office rollout planning\nM/EX',840,651,365,53,14.67,{align:'right'});
  s.speakerNotes.text='A proposed office communications and rollout plan for Account Management '+data.version+'. The office identity and visual design follow the supplied Better Task rollout reference. Communications will determine all dates later. Confirm the role owners and replace the bracketed links and contacts before use. The local fixture screenshots use fictional data. Source design: Better-Task-Rollout-and-Communication-Plan.pptx, supplied by the user. Content sources: ARCHITECTURE.md, docs/BASELINE.md, docs/ACCESS-ADMINISTRATION.md, and docs/KNOWN-ISSUES.md.';
}

for(const d of data.slides){
  const s=slide(d,d.kind==='closing'?'BA Closing':'BA Content');
  if(d.kind==='overview'){
    lede(s,d.lede);
    const gap=28,cw=(W-gap*3)/4;
    for(let i=0;i<d.cards.length;i++){
      const c=d.cards[i],x=M+i*(cw+gap);
      rect(s,c.label+' panel',x,243,cw,390,C.panel,13);
      await icon(s,c.icon,x+26,268);
      text(s,c.label+' label',c.label,x+26,347,cw-52,25,14.67,{bold:true,color:C.muted});
      text(s,c.label+' title',c.title,x+26,382,cw-52,100,29.33,{head:true,bold:true,color:C.navy});
      text(s,c.label+' body',c.body,x+26,477,cw-52,140,20);
    }
  } else if(d.kind==='table'){
    lede(s,d.lede);
    const vals=[d.columns,...d.rows];
    const y=225,h=d.rows.length===4?402:420;
    const t=s.tables.add({rows:vals.length,columns:d.columns.length,left:M,top:y,width:W,height:h,columnWidths:d.widths,values:vals});
    t.borders.assign({fill:C.rule,width:1,style:'solid'});
    for(let r=0;r<vals.length;r++){
      t.rows[r].height=r===0?42:(h-42)/d.rows.length;
      for(let c=0;c<d.columns.length;c++){
        const cell=t.getCell(r,c);cell.fill=r===0?C.navy:(r%2===0?C.panel:C.white);
        cell.text.style={fontSize:r===0?18.5:19.5,typeface:'Calibri',color:r===0?C.white:C.ink,bold:r===0||c===0,verticalAlignment:'middle',insets:{top:8,bottom:8,left:10,right:10},wrap:'square',autoFit:'none'};
      }
    }
    tableSlides.push(p.slides.items.length);
  } else if(d.kind==='two'){
    lede(s,d.lede);const y=d.lede?228:166;
    const cw=(W-32)/2;
    d.panels.forEach((v,i)=>{
      const x=M+i*(cw+32);
      rect(s,v.title+' panel',x,y,cw,635-y,C.panel,13);
      text(s,v.title+' heading',v.title,x+26,y+23,cw-52,48,32,{head:true,bold:true,color:C.navy});
      paired(s,v.items,x+26,y+91,cw-52,(635-y-100)/3,{size:22});
    });
  } else if(d.kind==='screen'){
    paired(s,d.steps,M,160,355,96,{size:20.5});
    const file=path.join(source,'images',d.image),bytes=new Uint8Array(await fs.readFile(file));
    const buf=Buffer.from(bytes),iw=buf.readUInt32BE(16),ih=buf.readUInt32BE(20),scale=Math.min(755/iw,457/ih);
    const sw=iw*scale,sh=ih*scale,x=450+(755-sw)/2;
    s.images.add({blob:bytes,contentType:'image/png',alt:d.alt,position:{left:x,top:166,width:sw,height:sh},fit:'contain'});
    rect(s,'Screenshot frame',x,166,sw,sh,'none',0,C.rule);
    text(s,'Screenshot caption',d.caption,450,637,755,28,14.67,{color:C.muted});
  } else if(d.kind==='status'){
    lede(s,d.lede);const gap=28,cw=(W-gap*2)/3;
    d.cards.forEach((c,i)=>{
      const x=M+i*(cw+gap);rect(s,c.label+' panel',x,248,cw,285,C.panel,13);
      text(s,c.label,c.label,x+26,273,cw-52,28,17.33,{bold:true,color:c.color});
      text(s,c.title,c.title,x+26,319,cw-52,50,32,{head:true,bold:true,color:c.color});
      text(s,c.label+' guidance',c.body,x+26,391,cw-52,120,23);
    });
    rect(s,'SharePoint distinction',M,563,W,82,C.navy,13);
    text(s,'SharePoint distinction text',d.callout,M+25,582,W-50,53,22,{color:C.white});
  } else if(d.kind==='phases'){
    const step=W/6;
    rect(s,'Phase line',M+step/2,181,W-step,2,C.rule);
    d.phaseLabels.forEach((label,i)=>{
      const x=M+i*step,active=d.active.includes(i);
      rect(s,label+' point',x+step/2-9,173,18,18,active?C.navy:C.white,9,C.navy);
      text(s,label+' phase',label,x,205,step,30,20,{bold:active,color:active?C.navy:C.muted,align:'center'});
    });
    const cw=(W-56)/3;
    d.cards.forEach((c,i)=>{
      const x=M+i*(cw+28);rect(s,c.title+' panel',x,265,cw,380,C.panel,13);
      text(s,c.title,c.title,x+26,291,cw-52,47,34,{head:true,bold:true,color:C.navy});
      text(s,c.title+' actions',c.body,x+26,355,cw-52,137,23);
      text(s,c.title+' readiness',c.gate,x+26,511,cw-52,78,20,{bold:true,color:C.navy});
      text(s,c.title+' owner','Owner: '+c.owner,x+26,604,cw-52,35,17.33,{bold:true,color:C.muted});
    });
  } else if(d.kind==='email'){
    text(s,'Audience',d.label,M,157,W,28,15.33,{bold:true,color:C.muted});
    rect(s,'Email panel',M,206,W,431,C.panel,13);
    text(s,'Subject','Subject: '+d.subject,M+30,229,W-60,40,25.33,{bold:true,color:C.navy});
    text(s,'Email copy',d.paragraphs.join('\n\n'),M+30,291,W-60,328,22);
  } else if(d.kind==='training'){
    text(s,'Invitation subject','Subject: '+d.subject,M,158,W,40,24,{bold:true,color:C.navy});
    text(s,'Invitation copy',d.invitation,M,212,W,126,23);
    const cw=(W-32)/2;
    d.columns.forEach((v,i)=>{
      const x=M+i*(cw+32);rect(s,v.title+' panel',x,368,cw,273,C.panel,13);
      text(s,v.title,v.title,x+26,390,cw-52,47,32,{head:true,bold:true,color:C.navy});
      v.items.forEach((item,index)=>text(s,v.title+' action '+(index+1),item,x+26,455+index*43,cw-52,40,22));
    });
  } else if(d.kind==='faq'){
    const cw=(W-60)/2;
    d.items.forEach(([q,a],i)=>{
      const x=M+(i%2)*(cw+60),y=162+Math.floor(i/2)*164;
      text(s,q,[{runs:[{run:q,style:{bold:true,color:C.navy}}],spaceAfter:10},{runs:[a]}],x,y,cw,136,24);
      if(i<4)rect(s,q+' rule',x,y+143,cw,1,C.rule);
    });
  } else if(d.kind==='incident'){
    paired(s,d.steps,M,167,415,155,{size:23});
    rect(s,'Issue notice panel',510,166,712,471,C.panel,13);
    text(s,'Draft notice label','DRAFT SERVICE NOTICE',538,194,651,30,15.33,{bold:true,color:C.muted});
    text(s,'Issue subject',d.subject,538,240,651,70,30,{head:true,bold:true,color:C.navy});
    text(s,'Issue message',d.copy,538,331,651,280,25);
  } else if(d.kind==='closing'){
    d.items.forEach((item,index)=>text(s,'Release decision '+(index+1),item,M,207+index*66,720,62,26,{color:C.white}));
    rect(s,'Communications responsibility',824,205,398,410,'#263D58',13);
    text(s,'Schedule owner',d.boxTitle,850,235,346,90,34,{head:true,bold:true,color:C.white});
    text(s,'Schedule ownership',d.boxBody,850,348,346,245,24,{color:C.white});
  } else throw new Error('Unknown slide kind '+d.kind);
}

// Keep a text snapshot and full-size previews private, outside the delivered folder.
await fs.writeFile(path.join(scratch,'final-content.ndjson'),(await p.inspect({kind:'slide,textbox,table,notes',maxChars:250000})).ndjson);
const candidate=path.join(scratch,'candidate.pptx');
await (await PresentationFile.exportPptx(p)).save(candidate);
for(let i=0;i<p.slides.items.length;i++){
  const image=await p.export({slide:p.slides.items[i],format:'png',scale:1});
  await fs.writeFile(path.join(scratch,'rendered',`${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await image.arrayBuffer()));
  console.log('Rendered slide',i+1);
}
const final=process.env.DECK_OUT||path.join(root,'docs/rollout/Account-Management-Rollout-and-Communications-Plan.pptx');
const result=await finalizePresentation({
  workspaceDir:root,candidatePath:candidate,finalPath:final,pythonExecutable:python,
  integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...tableSlides.flatMap(n=>['--require-native-table-slide',String(n)])],
  requiredNativeTableOwnerSlides:tableSlides,
  fontPolicy:{basis:'reference',families:['Garamond','Calibri'],referencePath:template,referenceSha256:crypto.createHash('sha256').update(await fs.readFile(template)).digest('hex')},
  verifyArtifactToolImport:true,receiptPath:path.join(scratch,path.basename(final)+'.validation.json')
});
console.log('Finalized',p.slides.items.length,'slides:',final);
console.log(JSON.stringify(result));
