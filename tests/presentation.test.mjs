import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile,writeFile,mkdir} from 'node:fs/promises';
import JSZip from 'jszip';
import {PDFDocument} from 'pdf-lib';
import {readPresentation} from '../src/presentation.js';
import {convert} from '../src/engine.js';

const bytes=await readFile('dist/samples/sample-slides.pptx');
const sample=()=>new File([bytes],'slides.pptx');

test('PPTX reads actual presentation order, slide text and embedded images',async()=>{
 const original=await readPresentation(sample());
 assert.equal(original.slides.length,2);assert.match(original.slides[0].text.join('\n'),/Programming lab notes/);
 assert.equal(original.slides[1].images.length,1);assert.ok(original.slides[1].images[0].width>0);
 const zip=await JSZip.loadAsync(bytes);let xml=await zip.file('ppt/presentation.xml').async('string');
 const ids=xml.match(/<p:sldId\s[^>]+\/>/g);assert.equal(ids.length,2);
 xml=xml.replace(ids.join(''),ids.slice().reverse().join(''));zip.file('ppt/presentation.xml',xml);
 const reordered=await readPresentation(new File([await zip.generateAsync({type:'uint8array'})],'reordered.pptx'));
 assert.match(reordered.slides[0].text.join('\n'),/A picture for your handout/);
});

test('PPTX rejects malformed packages and XML entities and ignores external images',async()=>{
 await assert.rejects(()=>readPresentation(new File(['not zip'],'bad.pptx')),/readable/);
 await assert.rejects(()=>readPresentation(new File(['old'],'bad.ppt')),/supported/);
 const zip=await JSZip.loadAsync(bytes);zip.file('ppt/presentation.xml','<!DOCTYPE x [<!ENTITY e SYSTEM "https://example.invalid">]><x/>');
 await assert.rejects(async()=>readPresentation(new File([await zip.generateAsync({type:'uint8array'})],'bad.pptx')),/unsupported XML/);
 const linked=await JSZip.loadAsync(bytes);const path='ppt/slides/_rels/slide2.xml.rels';
 let rel=await linked.file(path).async('string');rel=rel.replace(/Target="\.\.\/media\/[^"]+"/,'Target="https://example.invalid/private.png" TargetMode="External"');linked.file(path,rel);
 const result=await readPresentation(new File([await linked.generateAsync({type:'uint8array'})],'external.pptx'));
 assert.equal(result.slides[1].images.length,0);assert.equal(result.skippedImages,1);
});

test('PPTX converts to editable Word and paginated PDF handouts',async()=>{
 const word=await convert('pptx-word',[sample()],{});
 const zip=await JSZip.loadAsync(await word.blob.arrayBuffer()),xml=await zip.file('word/document.xml').async('string');
 assert.match(xml,/Programming lab notes/);assert.equal((xml.match(/<w:sectPr>/g)||[]).length,2);assert.match(xml,/<w:drawing>/);
 assert.match(word.note,/Original slide layouts/);assert.equal(word.name,'slides-handout.docx');
 const pdf=await convert('pptx-pdf',[sample()],{paper:'LETTER'});
 const doc=await PDFDocument.load(await pdf.blob.arrayBuffer());assert.equal(doc.getPageCount(),2);assert.equal(doc.getPage(0).getWidth(),612);
 await mkdir('qa',{recursive:true});await writeFile('qa/slides-handout.docx',new Uint8Array(await word.blob.arrayBuffer()));await writeFile('qa/slides-handout.pdf',new Uint8Array(await pdf.blob.arrayBuffer()));
});
