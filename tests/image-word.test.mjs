import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile,mkdir,writeFile} from 'node:fs/promises';
import JSZip from 'jszip';
import {imageDocument,fitImage,imagesToWord} from '../src/image-word.js';
import {convert} from '../src/engine.js';

test('image fitting preserves proportions and stays within portrait and landscape pages',()=>{
  for(const orientation of ['portrait','landscape'])for(const paper of ['A4','LETTER'])for(const [w,h] of [[4000,500],[500,4000],[100,100]]){
    const fit=fitImage(w,h,{orientation,paper});
    assert.ok(fit.width<=(fit.page[0]-84)*4/3);
    assert.ok(fit.height<=(fit.page[1]-144)*4/3);
    assert.ok(Math.abs(fit.width/fit.height-w/h)<.05);
    assert.ok(fit.width<=w&&fit.height<=h);
  }
  assert.throws(()=>fitImage(0,10),/dimensions/);
});

test('images are embedded in order with separate sections and optional captions',async()=>{
  const data=new Uint8Array(await readFile('dist/samples/sample-photo.jpg'));
  const images=[{name:'First & photo.jpg',data,type:'jpg',width:1200,height:800},{name:'Second photo.jpg',data,type:'jpg',width:800,height:1200}];
  const blob=await imageDocument(images,{paper:'LETTER',orientation:'landscape'});
  const zip=await JSZip.loadAsync(await blob.arrayBuffer());const xml=await zip.file('word/document.xml').async('string');
  assert.equal((xml.match(/<w:drawing>/g)||[]).length,2);
  assert.equal((xml.match(/<w:sectPr>/g)||[]).length,2);
  assert.match(xml,/w:orient="landscape"/);assert.match(xml,/w:w="15840"/);
  assert.ok(xml.indexOf('First &amp; photo.jpg')<xml.indexOf('Second photo.jpg'));
  assert.ok(Object.keys(zip.files).some(path=>path.startsWith('word/media/')&&path.endsWith('.jpg')));
  const noCaptions=await imageDocument(images,{imageCaptions:'no'});
  const noCaptionXml=await (await JSZip.loadAsync(await noCaptions.arrayBuffer())).file('word/document.xml').async('string');
  assert.doesNotMatch(noCaptionXml,/<w:t[^>]*>First/);
  await mkdir('qa',{recursive:true});await writeFile('qa/image-document.docx',new Uint8Array(await blob.arrayBuffer()));
  await assert.rejects(()=>imagesToWord([new File(['x'],'not-image.svg')]),/supported|Choose/i);
});

test('Apple Pages export is a real editable DOCX with an honest format label',async()=>{
  const result=await convert('code-pages',[new File(['int main() {\n    return 0;\n}\n'],'lab.c')]);
  assert.equal(result.name,'lab.docx');assert.match(result.note,/not a native .pages/);
  const zip=await JSZip.loadAsync(await result.blob.arrayBuffer());
  assert.match(await zip.file('word/document.xml').async('string'),/return 0;/);
});
