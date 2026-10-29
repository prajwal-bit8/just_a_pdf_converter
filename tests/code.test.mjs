import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile,mkdir,writeFile} from 'node:fs/promises';
import {PDFDocument} from 'pdf-lib';
import JSZip from 'jszip';
import {CODE_EXTENSIONS,readCodeFile,codeOptions,expandTabs,wrapCode,prepareCode,missingPdfGlyph} from '../src/code.js';
import {codeToPdf,codeToWord} from '../src/code-export.js';
import {convert} from '../src/engine.js';

const source=(text,name='main.c')=>new File([text],name,{type:'text/plain'});
const font=new Uint8Array(await readFile('dist/fonts/DejaVuSansMono.ttf'));

test('code reading accepts source extensions, preserves whitespace, and normalizes line endings',async()=>{
 for(const ext of ['c','cpp','java','py','js','ts','cs','v','ino','sql','txt']) {
  assert.ok(CODE_EXTENSIONS.includes(ext));
  assert.equal(await readCodeFile(source('int main() {\r\n\treturn 0;\r\n}\r\n','FILE.'+ext.toUpperCase())), 'int main() {\n\treturn 0;\n}\n');
 }
 const utf16=new Uint8Array([255,254,...Buffer.from('Hello Ω\n','utf16le')]);
 assert.equal(await readCodeFile(new File([utf16],'unicode.py')), 'Hello Ω\n');
});

test('code rejects binary, malformed encoding, empty and oversized input',async()=>{
 await assert.rejects(()=>readCodeFile(source('','empty.c')),/empty/);
 await assert.rejects(()=>readCodeFile(source(' \n\t')),/no code/);
 await assert.rejects(()=>readCodeFile(source('MZ\x00\x01')),/binary/);
 await assert.rejects(()=>readCodeFile(new File([new Uint8Array([0xc3,0x28])],'bad.c')),/UTF-8/);
 await assert.rejects(()=>readCodeFile(source('x'.repeat(1024*1024+1))),/1 MB/);
 await assert.rejects(()=>readCodeFile(source('x','program.exe')),/source-code/);
 await assert.rejects(()=>prepareCode(Array(11).fill(source('x'))),/1 and 10/);
 await assert.rejects(()=>prepareCode([source('x\n'.repeat(5001))]),/5,000/);
});

test('tab stops, blank lines, wrapping and continuation numbers preserve the entire source',()=>{
 assert.equal(expandTabs('a\tb\t',4),'a   b   ');
 const text='\tif (a < b) {\n\n'+ 'x'.repeat(700)+'\n}';
 for(const lineNumbers of ['yes','no']) {
  const opt=codeOptions({lineNumbers});const {rows}=wrapCode(text,opt);
  assert.equal(rows[0].text,'    if (a < b) {');assert.equal(rows[1].text,'');
  assert.equal(rows.filter(r=>r.number==='3').length,1);
  assert.equal(rows.slice(2,-1).map(r=>r.text).join(''),'x'.repeat(700));
 }
 assert.ok(wrapCode('x'.repeat(700),codeOptions({orientation:'landscape'})).rows.length<wrapCode('x'.repeat(700),codeOptions()).rows.length);
});

test('PDF font detects unsupported glyphs before export',async()=>{
 assert.equal(missingPdfGlyph(font,'C++ <stdio.h> Ω é\t\n'),undefined);
 assert.ok(missingPdfGlyph(font,'ಕನ್ನಡ'));
 await assert.rejects(()=>codeToPdf([source('// ಕನ್ನಡ')],{},()=>{},font),/Code to Word/);
});

test('code PDFs paginate, preserve file order, use landscape Letter, and export embedded fonts',async()=>{
 const a=source('#include <stdio.h>\n\nint main(void) {\n\tprintf("Hello Ω <code> & friends");\n\treturn 0;\n}\n');
 const b=source(Array.from({length:140},(_,i)=>`// Java line ${i+1}: ${'long_code_'.repeat(14)}`).join('\n'),'Main.java');
 const result=await codeToPdf([a,b],{codeTitle:'Lab 04 — Code record',paper:'LETTER',orientation:'landscape'},()=>{},font);
 assert.equal(result.name,'code-listing.pdf');const doc=await PDFDocument.load(await result.blob.arrayBuffer());
 assert.ok(doc.getPageCount()>3);assert.equal(doc.getPage(0).getWidth(),792);assert.equal(doc.getPage(0).getHeight(),612);
 assert.equal(doc.getTitle(),'Lab 04 — Code record');
 await mkdir('qa',{recursive:true});await writeFile('qa/code-listing.pdf',new Uint8Array(await result.blob.arrayBuffer()));
 const plain=await codeToPdf([a],{lineNumbers:'no'},()=>{},font);
 assert.equal((await PDFDocument.load(await plain.blob.arrayBuffer())).getPageCount(),1);
 await writeFile('qa/code-no-numbers.pdf',new Uint8Array(await plain.blob.arrayBuffer()));
});

test('Word keeps code editable and escapes markup without interpreting it',async()=>{
 const files=[source('if (a < b && b > 0) {\n\t// ಕನ್ನಡ Ω\n\n}\n'),source('<script>alert("literal code")</script>','index.html')];
 const result=await convert('code-word',files,{lineNumbers:'no',orientation:'landscape'});
 const zip=await JSZip.loadAsync(await result.blob.arrayBuffer());const xml=await zip.file('word/document.xml').async('string');
 assert.match(xml,/a &lt; b &amp;&amp; b &gt; 0/);assert.match(xml,/ {4}\/\/ ಕನ್ನಡ Ω/);
 assert.match(xml,/&lt;script&gt;alert/);assert.match(await zip.file('word/styles.xml').async('string'),/Courier New/);
 assert.match(xml,/w:orient="landscape"/);
 assert.equal((xml.match(/<w:sectPr>/g)||[]).length,2);assert.ok(xml.indexOf('main.c')<xml.indexOf('index.html'));
 assert.doesNotMatch(xml,/<w:hyperlink/);
 const numbered=await codeToWord([source('one\ntwo')],{lineNumbers:'yes'});
 const numberedXml=await (await JSZip.loadAsync(await numbered.blob.arrayBuffer())).file('word/document.xml').async('string');
 assert.match(numberedXml,/1 {3}/);assert.match(numberedXml,/2 {3}/);
 await mkdir('qa',{recursive:true});await writeFile('qa/code-listing.docx',new Uint8Array(await result.blob.arrayBuffer()));
});
