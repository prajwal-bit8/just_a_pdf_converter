import { PDFDocument, StandardFonts, rgb, degrees } from 'pdf-lib';
import { parseRanges, fileStem, pdfLines } from './core.js';

let pdfjsPromise;
export async function pdfjs() {
  if (!pdfjsPromise) pdfjsPromise = import('pdfjs-dist/build/pdf.mjs').then(lib => {
    lib.GlobalWorkerOptions.workerSrc = '/assets/pdf.worker.min.mjs';
    return lib;
  });
  return pdfjsPromise;
}
export async function openPdf(file) {
  const lib = await pdfjs();
  return lib.getDocument({data:new Uint8Array(await file.arrayBuffer()),standardFontDataUrl:'/assets/standard_fonts/',cMapUrl:'/assets/cmaps/',cMapPacked:true,wasmUrl:'/assets/wasm/',isEvalSupported:false}).promise;
}
export async function loadPdf(file) {
  try { return await PDFDocument.load(await file.arrayBuffer()); }
  catch (error) {
    if (/encrypt|password/i.test(error.message)) throw new Error('This PDF is password protected. Open it in your PDF app, save an unlocked copy, and try that copy.');
    throw new Error('This PDF could not be read. It may be damaged or unsupported. Try saving a new copy.');
  }
}
export async function previewPdf(file, withThumbnails = false, progress=()=>{}) {
  let doc;
  try {
    doc = await openPdf(file);
    if (withThumbnails && doc.numPages > 200) throw new Error('The page organizer supports up to 200 pages per file. Use Split PDF first.');
    const pages = [];
    if (withThumbnails) for (let i=1; i<=doc.numPages; i++) {
      const page = await doc.getPage(i), viewport=page.getViewport({scale:1});
      const vp=page.getViewport({scale:170/Math.max(viewport.width,viewport.height)});
      const canvas=document.createElement('canvas'); canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);
      await page.render({canvasContext:canvas.getContext('2d'),viewport:vp}).promise;
      pages.push({index:i-1,rotation:0,deleted:false,thumbnail:canvas.toDataURL('image/jpeg',.7)});
      canvas.width=0;canvas.height=0;page.cleanup();progress(i,doc.numPages);
    }
    return {count:doc.numPages,pages};
  } catch(error) {
    if (/password|encrypt/i.test(error.message)) throw new Error('Password-protected PDFs are not supported. Save an unlocked copy first.');
    throw error;
  } finally { if(doc) await doc.destroy(); }
}
const pdfBlob=bytes=>new Blob([bytes],{type:'application/pdf'});
const tick=()=>new Promise(resolve=>setTimeout(resolve,0));
async function renderPage(page,scale=1.5,quality=.8) {
  const base=page.getViewport({scale:1});
  const safeScale=Math.min(scale,4096/Math.max(base.width,base.height));
  const vp=page.getViewport({scale:safeScale});
  const canvas=document.createElement('canvas');canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);
  const context=canvas.getContext('2d');
  await page.render({canvasContext:context,viewport:vp,background:'rgb(255,255,255)'}).promise;
  const blob=await new Promise(r=>canvas.toBlob(r,'image/jpeg',quality));
  canvas.width=0;canvas.height=0;
  if(!blob) throw new Error('This page is too large to render in your browser. Try a smaller PDF.');
  return {blob,width:base.width,height:base.height};
}
async function zipResults(entries,name,progress) {
  const {default:JSZip}=await import('jszip');const zip=new JSZip();
  entries.forEach(({name,blob})=>zip.file(name,blob));
  progress(90,'Packaging your files…');
  return {name,blob:await zip.generateAsync({type:'blob',compression:'DEFLATE'})};
}
export async function extractLegacyDoc(buffer) {
  const CFB=await import('cfb');
  const container=CFB.read(new Uint8Array(buffer),{type:'array'});
  const stream=CFB.find(container,'WordDocument');
  if(!stream) throw new Error('This is not a supported Word 97–2003 document. Save it as .docx in Word and try again.');
  const word=new Uint8Array(stream.content), view=new DataView(word.buffer,word.byteOffset,word.byteLength);
  if(word.length<154 || view.getUint16(0,true)!==0xa5ec) throw new Error('The legacy Word file is damaged or unsupported.');
  const flags=view.getUint16(10,true);
  if(flags&0x8100) throw new Error('Password-protected Word documents are not supported. Save an unprotected copy first.');
  if(view.getUint16(2,true)<0x00c1) throw new Error('This older .doc format is unsupported. Save it as .docx in Word first.');
  const tableEntry=CFB.find(container,(flags&0x0200)?'1Table':'0Table');
  if(!tableEntry) throw new Error('The Word file is missing its text table. Save it as .docx and try again.');
  const table=new Uint8Array(tableEntry.content),tv=new DataView(table.buffer,table.byteOffset,table.byteLength);
  const csw=view.getUint16(32,true), cslwPos=34+csw*2, cslw=view.getUint16(cslwPos,true), lwPos=cslwPos+2;
  const mainCharCount=view.getUint32(lwPos+12,true), pairsPos=lwPos+cslw*4+2;
  const clxOffset=view.getUint32(pairsPos+33*8,true),clxLength=view.getUint32(pairsPos+33*8+4,true);
  if(!clxLength || clxOffset+clxLength>table.length) throw new Error('This .doc text structure is unsupported. Save it as .docx first.');
  let pos=clxOffset;
  while(pos<table.length && table[pos]===1) pos+=3+tv.getUint16(pos+1,true);
  if(table[pos]!==2) throw new Error('Could not read the Word text table. Save the document as .docx.');
  const length=tv.getUint32(pos+1,true),start=pos+5,n=(length-4)/12;
  if(!Number.isInteger(n)||n<1||start+length>table.length) throw new Error('The Word text table is damaged.');
  let result='';
  for(let i=0;i<n;i++) {
    const cp=tv.getUint32(start+i*4,true),next=tv.getUint32(start+(i+1)*4,true);
    if(cp>=mainCharCount) break;
    const count=Math.min(next,mainCharCount)-cp,pcd=start+(n+1)*4+i*8,raw=tv.getUint32(pcd+2,true);
    const compressed=!!(raw&0x40000000),offset=(raw&0x3fffffff)/(compressed?2:1),bytes=count*(compressed?1:2);
    if(offset+bytes>word.length) throw new Error('The Word document contains an invalid text reference.');
    result+=new TextDecoder(compressed?'windows-1252':'utf-16le').decode(word.slice(offset,offset+bytes));
  }
  // Keep field results, discard field instructions, and normalize Word paragraph/cell markers.
  let cleaned='',stack=[];
  for(const c of result) {
    if(c==='\x13'){stack.push(false);continue;}if(c==='\x14'){if(stack.length)stack[stack.length-1]=true;continue;}if(c==='\x15'){stack.pop();continue;}
    if(stack.some(v=>!v))continue;cleaned+=c;
  }
  return cleaned.replace(/\r/g,'\n').replace(/\x07/g,'\t').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').trim();
}
async function wordToPdf(file,opt,progress) {
  const [{default:pdfMake},{default:fonts}]=await Promise.all([import('pdfmake/build/pdfmake.js'),import('pdfmake/build/vfs_fonts.js')]);
  pdfMake.addVirtualFileSystem(fonts);
  progress(20,'Reading your Word document…');
  let content,note;
  if(file.name.toLowerCase().endsWith('.doc')) {
    const text=await extractLegacyDoc(await file.arrayBuffer());
    if(!text) throw new Error('No readable text was found in this .doc file. Save it as .docx to convert it.');
    content=text.split('\n').map(text=>({text:text||' ',margin:[0,0,0,5]}));
    note='Legacy .doc conversion includes document text only. Images and original formatting are not retained.';
  } else {
    const [{default:mammoth},{default:htmlToPdfmake},{default:DOMPurify}]=await Promise.all([import('mammoth/mammoth.browser.js'),import('html-to-pdfmake'),import('dompurify')]);
    const result=await mammoth.convertToHtml({arrayBuffer:await file.arrayBuffer()},{externalFileAccess:false});
    const clean=DOMPurify.sanitize(result.value,{FORBID_TAGS:['script','iframe','object','embed','style','link','form'],FORBID_ATTR:['srcset']});
    const div=document.createElement('div');div.innerHTML=clean;
    div.querySelectorAll('img').forEach(img=>{if(!/^data:image\/(png|jpeg);base64,/i.test(img.getAttribute('src')||'')) img.remove(); else {img.style.maxWidth='450px';img.style.height='auto';img.setAttribute('width','450');}});
    div.querySelectorAll('a').forEach(a=>{if(!/^https?:/i.test(a.getAttribute('href')||''))a.removeAttribute('href');});
    if(!div.textContent.trim()&&!div.querySelector('img'))throw new Error('No convertible content was found in this Word document.');
    content=htmlToPdfmake(div.innerHTML,{window,tableAutoSize:true,defaultStyles:{p:{margin:[0,0,0,8]},h1:{fontSize:22,bold:true,margin:[0,12,0,10]},h2:{fontSize:17,bold:true,margin:[0,10,0,8]},img:{margin:[0,8,0,8]}}});
    note='Headings, text, tables and supported images are rebuilt. Complex layouts, headers and custom fonts may differ.';
  }
  progress(62,'Creating your PDF…');await tick();
  const def={pageSize:opt.paper||'A4',pageMargins:[40,40,40,40],defaultStyle:{font:'Roboto',fontSize:11,lineHeight:1.25},content,info:{title:fileStem(file.name),creator:'Just a File Converter'}};
  const blob=await new Promise((resolve,reject)=>{
    try {
      const stream=pdfMake.createPdf(def).getStream(), chunks=[];
      stream.on('data',chunk=>chunks.push(new Uint8Array(chunk)));
      stream.on('error',reject);
      stream.on('end',()=>resolve(new Blob(chunks,{type:'application/pdf'})));
      stream.end();
    }catch(error){reject(error);}
  });
  return {name:fileStem(file.name)+'.pdf',blob,note};
}
async function pdfToWord(file,opt,progress) {
  const {Document,Packer,Paragraph,TextRun,ImageRun}=await import('docx');
  let doc;
  try {
    doc=await openPdf(file);if(doc.numPages>200)throw new Error('Word conversion supports up to 200 pages. Split your PDF into smaller parts first.');
    const sections=[];let textLength=0,emptyPages=0;
    for(let i=1;i<=doc.numPages;i++) {
      const page=await doc.getPage(i),vp=page.getViewport({scale:1});let children=[];
      if(opt.wordMode==='images') {
        const rendered=await renderPage(page,1.5,.87);
        children=[new Paragraph({children:[new ImageRun({data:new Uint8Array(await rendered.blob.arrayBuffer()),type:'jpg',transformation:{width:Math.min(624,vp.width),height:Math.min(624,vp.width)*vp.height/vp.width}})]})];
      } else {
        const text=await page.getTextContent(),lines=pdfLines(text.items);textLength+=lines.reduce((n,l)=>n+l.text.length,0);
        if(!lines.length){emptyPages++;children=[new Paragraph({children:[new TextRun({text:'[No selectable text on this page]',italics:true,color:'999999'})]})];}
        else children=lines.map(line=>new Paragraph({spacing:{after:100},children:[new TextRun({text:line.text,size:Math.round(Math.max(9,Math.min(28,line.size))*2)})]}));
      }
      sections.push({properties:{page:{size:{width:Math.round(vp.width*20),height:Math.round(vp.height*20)},margin:{top:600,bottom:600,left:600,right:600}}},children});
      page.cleanup();progress(10+Math.round(i/doc.numPages*75),`Converting page ${i} of ${doc.numPages}…`);
    }
    if(opt.wordMode!=='images'&&!textLength)throw new Error('This PDF has no selectable text. Choose “Page images” to preserve its pages in Word. Editable scanned text requires OCR, which is not included.');
    const output=new Document({creator:'Just a File Converter',title:fileStem(file.name),sections});
    const note=opt.wordMode==='images'?'Each PDF page is an image in Word. Its text is not editable.':`Text is editable. Original columns, tables, images and exact spacing are not reconstructed.${emptyPages?` ${emptyPages} page(s) had no selectable text and are marked in the document.`:''}`;
    return {name:fileStem(file.name)+'.docx',blob:await Packer.toBlob(output),note};
  } finally{if(doc)await doc.destroy();}
}
export async function convert(tool,files,opt={},progress=()=>{}) {
  progress(5,'Preparing your files…');await tick();
  const stem=fileStem(files[0].name);
  if(tool==='code-pdf'||tool==='code-word'||tool==='code-pages') {
    const code=await import('./code-export.js');
    if(tool==='code-pdf')return code.codeToPdf(files,opt,progress);
    const result=await code.codeToWord(files,opt,progress);
    if(tool==='code-pages')result.note='Download this .docx, then open it with Apple Pages on your Mac to edit or save as .pages. This is a Pages-compatible Word document, not a native .pages file. Code is not executed.';
    return result;
  }
  if(tool==='images-word')return (await import('./image-word.js')).imagesToWord(files,opt,progress);
  if(tool==='pptx-word'||tool==='pptx-pdf')return (await import('./presentation.js')).presentationExport(files[0],tool==='pptx-word'?'word':'pdf',opt,progress);
  if(tool==='word-pdf')return wordToPdf(files[0],opt,progress);
  if(tool==='pdf-word')return pdfToWord(files[0],opt,progress);
  if(tool==='merge') {
    if(files.length<2)throw new Error('Add at least two PDFs to merge.');
    const output=await PDFDocument.create();
    for(let i=0;i<files.length;i++){const doc=await loadPdf(files[i]);const pages=await output.copyPages(doc,doc.getPageIndices());pages.forEach(p=>output.addPage(p));progress(10+(i+1)/files.length*75,`Merging file ${i+1} of ${files.length}…`);await tick();}
    return {name:'merged-document.pdf',blob:pdfBlob(await output.save()),note:`${files.length} PDFs combined in the order shown.`};
  }
  if(tool==='images-pdf') {
    const output=await PDFDocument.create();
    for(let i=0;i<files.length;i++){
      const bytes=await files[i].arrayBuffer();let image;
      try { image=/\.png$/i.test(files[i].name)?await output.embedPng(bytes):await output.embedJpg(bytes); }catch{throw new Error(`“${files[i].name}” is not a readable JPG or PNG image.`);}
      const natural=opt.imagePaper==='original',land=opt.orientation==='landscape';
      const size=natural?[image.width,image.height]:land?[841.89,595.28]:[595.28,841.89];
      const page=output.addPage(size),margin=natural?0:24,scale=Math.min((size[0]-margin*2)/image.width,(size[1]-margin*2)/image.height);
      page.drawImage(image,{x:(size[0]-image.width*scale)/2,y:(size[1]-image.height*scale)/2,width:image.width*scale,height:image.height*scale});
      progress(10+(i+1)/files.length*75,`Adding image ${i+1} of ${files.length}…`);await tick();
    }
    return{name:stem+'.pdf',blob:pdfBlob(await output.save()),note:'One image per page, in the order shown.'};
  }
  if(['pdf-jpg','compress','text'].includes(tool)) {
    let doc;
    try {
      doc=await openPdf(files[0]);
      const selected=parseRanges(opt.range||'',doc.numPages);
      if(tool!=='text'&&selected.length>100)throw new Error('Image rendering supports up to 100 pages at a time. Enter a smaller page range or split the PDF first.');
      const entries=[],textPages=[];const output=tool==='compress'?await PDFDocument.create():null;
      for(let j=0;j<selected.length;j++){
        const i=selected[j],page=await doc.getPage(i+1);
        if(tool==='text')textPages.push(`--- Page ${i+1} ---\n${pdfLines((await page.getTextContent()).items).map(l=>l.text).join('\n')}`);
        else {
          const level=opt.quality||'balanced',scale=tool==='compress'?(level==='small'?.85:level==='high'?1.7:1.2):(opt.resolution==='high'?2:1.3),quality=tool==='compress'?(level==='small'?.45:level==='high'?.86:.65):.9;
          const r=await renderPage(page,scale,quality);
          if(tool==='compress'){const image=await output.embedJpg(await r.blob.arrayBuffer());const p=output.addPage([r.width,r.height]);p.drawImage(image,{x:0,y:0,width:r.width,height:r.height});}
          else entries.push({name:`${stem}-page-${i+1}.jpg`,blob:r.blob});
        }
        page.cleanup();progress(10+(j+1)/selected.length*75,`Processing page ${j+1} of ${selected.length}…`);await tick();
      }
      if(tool==='text'){
        if(!textPages.some(s=>s.replace(/--- Page \d+ ---/,'').trim()))throw new Error('No selectable text was found. Scanned PDFs need OCR, which is not included.');
        return{name:stem+'.txt',blob:new Blob([textPages.join('\n\n')],{type:'text/plain;charset=utf-8'}),note:'Selectable text extracted. Scanned text and images are not included.'};
      }
      if(tool==='pdf-jpg')return entries.length===1?{...entries[0],note:'Your PDF page, exported as a JPG.'}:{...await zipResults(entries,stem+'-images.zip',progress),note:`${entries.length} JPG images in one ZIP file.`};
      const blob=pdfBlob(await output.save());
      if(blob.size>=files[0].size)return{name:stem+'-original.pdf',blob:files[0],note:'This PDF is already smaller than the compressed result. Your original file has been kept unchanged.'};
      const reduction=Math.round((1-blob.size/files[0].size)*100);
      return{name:stem+'-compressed.pdf',blob,note:`${reduction}% smaller. Pages are flattened to images; searchable text, links and forms are removed.`};
    }finally{if(doc)await doc.destroy();}
  }
  const doc=await loadPdf(files[0]),count=doc.getPageCount();
  if(tool==='split') {
    const pages=parseRanges(opt.range||'',count);
    if(opt.splitMode==='each'){
      if(pages.length>300)throw new Error('Split up to 300 pages at once by entering a page range.');
      const entries=[];
      for(let j=0;j<pages.length;j++){const output=await PDFDocument.create();(await output.copyPages(doc,[pages[j]])).forEach(p=>output.addPage(p));entries.push({name:`${stem}-page-${pages[j]+1}.pdf`,blob:pdfBlob(await output.save())});progress(10+(j+1)/pages.length*75,`Extracting page ${j+1} of ${pages.length}…`);await tick();}
      return entries.length===1?entries[0]:{...await zipResults(entries,stem+'-split.zip',progress),note:`${entries.length} individual PDFs in a ZIP file.`};
    }
    const output=await PDFDocument.create();(await output.copyPages(doc,pages)).forEach(p=>output.addPage(p));
    return{name:stem+'-extracted.pdf',blob:pdfBlob(await output.save()),note:`${pages.length} selected pages extracted.`};
  }
  if(tool==='organize') {
    const selected=(opt.pages||[]).filter(p=>!p.deleted);if(!selected.length)throw new Error('Keep at least one page in your PDF.');
    const output=await PDFDocument.create();
    for(const p of selected){const [copy]=await output.copyPages(doc,[p.index]);copy.setRotation(degrees((copy.getRotation().angle+p.rotation)%360));output.addPage(copy);}
    return{name:stem+'-organized.pdf',blob:pdfBlob(await output.save()),note:`${selected.length} pages saved in your chosen order.`};
  }
  const pages=parseRanges(opt.range||'',count);
  if(tool==='rotate') {for(const i of pages){const p=doc.getPage(i);p.setRotation(degrees((p.getRotation().angle+Number(opt.angle||90))%360));}}
  if(tool==='watermark') {
    const text=String(opt.watermark||'').trim();if(!text)throw new Error('Enter the text for your watermark.');if(text.length>80)throw new Error('Keep the watermark under 80 characters.');
    const font=await doc.embedFont(StandardFonts.HelveticaBold);
    try {font.encodeText(text);}catch{throw new Error('Watermarks support Latin text in this version. Use letters supported by the standard PDF font.');}
    for(const i of pages){const p=doc.getPage(i),w=p.getWidth(),h=p.getHeight(),size=Math.min(48,w*.7/Math.max(font.widthOfTextAtSize(text,1),1));const width=font.widthOfTextAtSize(text,size);p.drawText(text,{x:(w-width)/2,y:h/2,size,font,color:rgb(.42,.34,.56),opacity:Number(opt.opacity||.2),rotate:degrees(0)});}
  }
  if(tool==='numbers') {
    const font=await doc.embedFont(StandardFonts.Helvetica),start=Number(opt.start||1);if(!Number.isInteger(start)||start<1)throw new Error('Starting number must be a positive whole number.');
    pages.forEach((idx,n)=>{const p=doc.getPage(idx),text=opt.numberFormat==='total'?`${n+start} / ${pages.length+start-1}`:String(n+start),width=font.widthOfTextAtSize(text,10);p.drawText(text,{x:(p.getWidth()-width)/2,y:opt.position==='top'?p.getHeight()-24:20,size:10,font,color:rgb(.4,.36,.45)});});
  }
  if(!['rotate','watermark','numbers'].includes(tool))throw new Error('Choose a supported tool.');
  progress(85,'Saving your PDF…');
  return{name:stem+({rotate:'-rotated',watermark:'-watermarked',numbers:'-numbered'}[tool])+'.pdf',blob:pdfBlob(await doc.save()),note:`Updated ${pages.length} page${pages.length===1?'':'s'}.`};
}
