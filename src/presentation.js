import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {PDFDocument} from 'pdf-lib';
import {fileStem,validateFiles} from './core.js';
import {missingPdfGlyph} from './code.js';

const descendants=(node,name)=>Array.from(node.getElementsByTagNameNS('*',name));
const relAttr=(node,name)=>Array.from(node.attributes||[]).find(a=>a.localName===name&&a.prefix)?.value;
function parseXml(text) {
  if(text.length>4*1024*1024||/<!DOCTYPE|<!ENTITY/i.test(text))throw new Error('This presentation contains unsupported XML.');
  let invalid=false;
  const doc=new DOMParser({errorHandler:{warning:()=>{},error:()=>{invalid=true;},fatalError:()=>{invalid=true;}}}).parseFromString(text,'application/xml');
  if(invalid||!doc?.documentElement)throw new Error('This presentation contains damaged XML.');
  return doc;
}
function resolvePart(base,target) {
  if(!target||/^[a-z][a-z\d+.-]*:|^\/\/|\\/i.test(target))return null;
  const parts=target.startsWith('/')?[]:base.split('/').slice(0,-1);
  for(const part of target.split('/')){if(part==='..'){if(!parts.length)return null;parts.pop();}else if(part&&part!=='.')parts.push(part);}
  return parts.join('/');
}
function relationships(doc,base) {
  return new Map(descendants(doc,'Relationship').filter(r=>r.getAttribute('TargetMode')!=='External').map(r=>[r.getAttribute('Id'),{path:resolvePart(base,r.getAttribute('Target')),type:r.getAttribute('Type')}]));
}
function paragraphText(node) {
  let text='';
  function visit(n){if(n.localName==='t')text+=n.textContent;else if(n.localName==='br')text+='\n';else for(const c of Array.from(n.childNodes||[]))visit(c);}
  visit(node);
  const bullet=descendants(node,'buChar').length||descendants(node,'buAutoNum').length;
  return (bullet?'• ':'')+text;
}
const asBase64=data=>{let s='';for(let i=0;i<data.length;i+=8192)s+=String.fromCharCode(...data.subarray(i,i+8192));return btoa(s);};

export async function readPresentation(file,progress=()=>{}) {
  validateFiles([file],['pptx']);
  let zip;try{zip=await JSZip.loadAsync(await file.arrayBuffer());}catch{throw new Error('This is not a readable .pptx. Password-protected and older .ppt files are not supported.');}
  const entries=Object.values(zip.files);
  if(entries.length>5000||entries.reduce((n,e)=>n+(e._data?.uncompressedSize||0),0)>200*1024*1024)throw new Error('This presentation is too large when unpacked. Use a smaller presentation.');
  async function xml(path,optional=false){const item=zip.file(path);if(!item){if(optional)return parseXml('<Relationships/>');throw new Error('The presentation is missing required slide information.');}return parseXml(await item.async('string'));}
  const presentation=await xml('ppt/presentation.xml'),rels=relationships(await xml('ppt/_rels/presentation.xml.rels'),'ppt/presentation.xml');
  const ids=descendants(presentation,'sldId');
  if(!ids.length||ids.length>100)throw new Error('Choose a presentation with 1 to 100 slides.');
  const slides=[],imageCache=new Map(),imageReader=await PDFDocument.create();let totalText=0,totalImages=0,skippedImages=0;
  for(let i=0;i<ids.length;i++) {
    const rel=rels.get(relAttr(ids[i],'id'));
    if(!rel?.path||!rel.type.endsWith('/slide'))throw new Error('The presentation contains a missing or unsupported slide.');
    const doc=await xml(rel.path),slash=rel.path.lastIndexOf('/'),relPath=rel.path.slice(0,slash+1)+'_rels/'+rel.path.slice(slash+1)+'.rels';
    const slideRels=relationships(await xml(relPath,true),rel.path);
    const text=descendants(doc,'p').map(paragraphText).filter(t=>t.trim());
    totalText+=text.join('').length;if(totalText>400000)throw new Error('Keep the presentation under 400,000 text characters.');
    const images=[];
    for(const blip of descendants(doc,'blip')) {
      const imageRel=slideRels.get(relAttr(blip,'embed'));
      if(!imageRel?.path||!imageRel.type.endsWith('/image')||!zip.file(imageRel.path)){skippedImages++;continue;}
      if(++totalImages>200)throw new Error('Keep the presentation under 200 images.');
      let image=imageCache.get(imageRel.path);
      if(image===undefined){
        image=null;
        const type=/\.png$/i.test(imageRel.path)?'png':/\.jpe?g$/i.test(imageRel.path)?'jpg':null;
        if(type)try{
          const data=await zip.file(imageRel.path).async('uint8array');
          const embedded=type==='png'?await imageReader.embedPng(data):await imageReader.embedJpg(data);
          if(embedded.width*embedded.height>40000000)throw new Error('too large');
          image={data,type,width:embedded.width,height:embedded.height,url:'data:image/'+(type==='jpg'?'jpeg':'png')+';base64,'+asBase64(data)};
        }catch{}
        imageCache.set(imageRel.path,image);
      }
      if(image)images.push(image);else skippedImages++;
    }
    slides.push({number:i+1,text,images});
    progress(10+(i+1)/ids.length*45,`Reading slide ${i+1} of ${ids.length}…`);
  }
  if(!slides.some(s=>s.text.length||s.images.length))throw new Error('No supported slide text or JPG/PNG images were found. Use PowerPoint to export this presentation.');
  return {slides,skippedImages};
}

export async function presentationExport(file,format,opt={},progress=()=>{}) {
  const {slides,skippedImages}=await readPresentation(file,progress),stem=fileStem(file.name);
  const page=opt.paper==='LETTER'?[612,792]:[595.28,841.89];if(opt.orientation==='landscape')page.reverse();
  const note=`Handout export: each slide starts on a new page. Text and embedded JPG/PNG images are included. Original slide layouts, backgrounds, charts, SmartArt, notes, animations and media are not reproduced. Tables become plain text.${skippedImages?' '+skippedImages+' unsupported or unavailable image(s) omitted.':''}`;
  progress(65,'Creating your slide handout…');
  let blob;
  if(format==='word') {
    const {Document,Packer,Paragraph,TextRun,ImageRun,SectionType,PageOrientation}=await import('docx');
    const sections=slides.map(slide=>({properties:{type:SectionType.NEXT_PAGE,page:{size:{width:Math.round(Math.min(...page)*20),height:Math.round(Math.max(...page)*20),orientation:opt.orientation==='landscape'?PageOrientation.LANDSCAPE:PageOrientation.PORTRAIT},margin:{top:840,bottom:840,left:840,right:840}}},children:[
      new Paragraph({keepNext:true,spacing:{after:240},children:[new TextRun({text:'Slide '+slide.number,bold:true,size:32,color:'493582'})]}),
      ...slide.text.map(text=>new Paragraph({spacing:{after:120},children:[new TextRun({text})]})),
      ...slide.images.map(img=>{const scale=Math.min((page[0]-84)*4/3/img.width,(page[1]-140)*4/3/img.height,1);return new Paragraph({spacing:{after:160},children:[new ImageRun({data:img.data,type:img.type,transformation:{width:Math.floor(img.width*scale),height:Math.floor(img.height*scale)}})]});}),
      ...(!slide.text.length&&!slide.images.length?[new Paragraph('No supported text or images on this slide.')]:[])
    ]}));
    blob=await Packer.toBlob(new Document({creator:'Just a File Converter',title:stem,styles:{default:{document:{run:{font:'Arial',size:22}}}},sections}));
  } else {
    const [{default:pdfMake},{default:vfs}]=await Promise.all([import('pdfmake/build/pdfmake.js'),import('pdfmake/build/vfs_fonts.js')]);
    const font=Uint8Array.from(atob(vfs['Roboto-Regular.ttf']),c=>c.charCodeAt(0));
    if(missingPdfGlyph(font,slides.flatMap(s=>s.text).join('')))throw new Error('Some slide characters are unavailable in the PDF font. Use PPTX to Word and choose a font that supports the text.');
    const content=slides.flatMap((slide,i)=>[
      {text:'Slide '+slide.number,fontSize:18,bold:true,color:'#493582',margin:[0,0,0,14],pageBreak:i?'before':undefined},
      ...slide.text.map(text=>({text,margin:[0,0,0,8]})),
      ...slide.images.map(img=>({image:img.url,fit:[page[0]-84,page[1]-150],margin:[0,8,0,12]})),
      ...(!slide.text.length&&!slide.images.length?[{text:'No supported text or images on this slide.'}]:[])
    ]);
    const definition={pageSize:opt.paper==='LETTER'?'LETTER':'A4',pageOrientation:opt.orientation==='landscape'?'landscape':'portrait',pageMargins:[42,42,42,42],defaultStyle:{font:'Roboto',fontSize:11,lineHeight:1.2},content,info:{title:stem,creator:'Just a File Converter'}};
    blob=await new Promise((resolve,reject)=>{try{const stream=pdfMake.createPdf(definition,null,{Roboto:{normal:'Roboto-Regular.ttf',bold:'Roboto-Medium.ttf'}},vfs).getStream(),chunks=[];stream.on('data',chunk=>chunks.push(new Uint8Array(chunk)));stream.on('error',reject);stream.on('end',()=>resolve(new Blob(chunks,{type:'application/pdf'})));stream.end();}catch(error){reject(error);}});
  }
  return {name:stem+'-handout.'+(format==='word'?'docx':'pdf'),blob,note};
}
