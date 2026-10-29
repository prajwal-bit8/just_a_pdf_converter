import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {PDFDocument} from 'pdf-lib';
import * as CFB from 'cfb';
import {parseRanges,validateFiles,pdfLines} from '../src/core.js';
import {convert,extractLegacyDoc} from '../src/engine.js';
const file=async name=>new File([await readFile('dist/samples/'+name)],name);
const load=async result=>PDFDocument.load(await result.blob.arrayBuffer());
test('page ranges respect order, deduplicate overlaps, and reject malformed or out-of-bounds input',()=>{
 assert.deepEqual(parseRanges('3, 1-2, 2',5),[2,0,1]);assert.deepEqual(parseRanges('',3),[0,1,2]);
 for(const value of ['0','6','4-2','1,,2','hello','1-99','-1','1.5'])assert.throws(()=>parseRanges(value,5));
});
test('file validation rejects unsupported, empty, oversized, and over-budget files',()=>{
 assert.throws(()=>validateFiles([{name:'thing.exe',size:1}],['pdf']));assert.throws(()=>validateFiles([{name:'empty.pdf',size:0}],['pdf']));assert.throws(()=>validateFiles([{name:'large.pdf',size:51*1024*1024}],['pdf']));assert.throws(()=>validateFiles([{name:'a.pdf',size:50*1024*1024}],['pdf'],[{size:110*1024*1024}]));
 validateFiles([{name:'REPORT.PDF',size:500}],['pdf']);
});
test('merge produces expected total pages in the chosen order',async()=>{
 const r=await convert('merge',[await file('meeting-notes.pdf'),await file('project-brief.pdf')]);const d=await load(r);assert.equal(d.getPageCount(),4);assert.equal(r.name,'merged-document.pdf');
 await assert.rejects(()=>convert('merge',[new File(['x'],'x.pdf')]),/at least two/);
});
test('split extracts only the requested pages and rejects invalid ranges',async()=>{
 const f=await file('project-brief.pdf');const r=await convert('split',[f],{range:'3,1',splitMode:'selected'});assert.equal((await load(r)).getPageCount(),2);
 await assert.rejects(()=>convert('split',[f],{range:'4'}),/between 1 and 3/);
});
test('organize deletes, reorders and rotates without changing the source',async()=>{
 const f=await file('project-brief.pdf');const r=await convert('organize',[f],{pages:[{index:2,rotation:90},{index:0,rotation:0},{index:1,deleted:true}]});
 const d=await load(r);assert.equal(d.getPageCount(),2);assert.equal(d.getPage(0).getRotation().angle,90);assert.equal((await PDFDocument.load(await f.arrayBuffer())).getPageCount(),3);
 await assert.rejects(()=>convert('organize',[f],{pages:[{index:0,deleted:true}]}),/at least one/);
});
test('rotation applies only to selected pages',async()=>{
 const r=await convert('rotate',[await file('project-brief.pdf')],{range:'2',angle:'270'});const d=await load(r);assert.deepEqual(d.getPages().map(p=>p.getRotation().angle),[0,270,0]);
});
test('page numbering and watermark create readable PDFs and validate inputs',async()=>{
 const f=await file('project-brief.pdf');assert.equal((await load(await convert('numbers',[f],{start:'5',position:'top',numberFormat:'total'}))).getPageCount(),3);
 assert.equal((await load(await convert('watermark',[f],{watermark:'DRAFT',opacity:'.3'}))).getPageCount(),3);
 await assert.rejects(()=>convert('watermark',[f],{watermark:''}),/Enter the text/);await assert.rejects(()=>convert('numbers',[f],{start:'-1'}),/positive whole/);
});
test('JPG to PDF creates one page per image',async()=>{
 const f=await file('sample-photo.jpg');const r=await convert('images-pdf',[f,f],{imagePaper:'a4',orientation:'landscape'});const d=await load(r);assert.equal(d.getPageCount(),2);assert.ok(d.getPage(0).getWidth()>d.getPage(0).getHeight());
});
test('line reconstruction respects page order and word spacing',()=>{
 const items=[{str:'World',width:30,transform:[1,0,0,12,80,700]},{str:'Second line',width:40,transform:[1,0,0,12,40,680]},{str:'Hello',width:30,transform:[1,0,0,12,40,700]}];assert.deepEqual(pdfLines(items).map(r=>r.text),['Hello World','Second line']);
});
function legacyFixture(text,unicode=false){
 const c=CFB.utils.cfb_new();const w=Buffer.alloc(4096),t=Buffer.alloc(4096);w.writeUInt16LE(0xa5ec,0);w.writeUInt16LE(0xc1,2);w.writeUInt16LE(0,10);w.writeUInt16LE(14,32);w.writeUInt16LE(22,62);w.writeUInt32LE(text.length,76);w.writeUInt16LE(93,152);w.writeUInt32LE(0,154+33*8);w.writeUInt32LE(21,154+33*8+4);Buffer.from(text,unicode?'utf16le':'latin1').copy(w,1024);t[0]=2;t.writeUInt32LE(16,1);t.writeUInt32LE(0,5);t.writeUInt32LE(text.length,9);t.writeUInt32LE(unicode?1024:(0x40000000+2048),15);CFB.utils.cfb_add(c,'WordDocument',w);CFB.utils.cfb_add(c,'0Table',t);return CFB.write(c,{type:'buffer'});
}
test('legacy DOC extracts ANSI and Unicode pieces and filters field instructions',async()=>{
 for(const unicode of [false,true]){const raw=legacyFixture('Hello Word\rSecond line\r',unicode);const array=raw.buffer.slice(raw.byteOffset,raw.byteOffset+raw.byteLength);assert.equal(await extractLegacyDoc(array),'Hello Word\nSecond line');}
 const raw=legacyFixture('A \x13HYPERLINK secret\x14Link text\x15 here');assert.equal(await extractLegacyDoc(raw.buffer.slice(raw.byteOffset,raw.byteOffset+raw.byteLength)),'A Link text here');
});
