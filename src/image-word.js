import {extension,fileStem,validateFiles} from './core.js';

export const IMAGE_WORD_EXTENSIONS=['jpg','jpeg','png','webp','gif','bmp'];
export function fitImage(width,height,opt={}) {
  if(!Number.isFinite(width)||!Number.isFinite(height)||width<=0||height<=0)throw new Error('This image has invalid dimensions.');
  const page=opt.paper==='LETTER'?[612,792]:[595.28,841.89];
  if(opt.orientation==='landscape')page.reverse();
  const maxWidth=(page[0]-84)*4/3,maxHeight=(page[1]-84-(opt.imageCaptions==='no'?18:60))*4/3;
  const scale=Math.min(1,maxWidth/width,maxHeight/height);
  return {width:Math.max(1,Math.floor(width*scale)),height:Math.max(1,Math.floor(height*scale)),page};
}

async function normalizeImage(file) {
  let bitmap,canvas;
  try {
    // Browser decoding applies photo orientation; rasterization also removes metadata.
    bitmap=await createImageBitmap(file);
    if(bitmap.width*bitmap.height>40000000)throw new Error('Choose images smaller than 40 megapixels.');
    const scale=Math.min(1,3000/Math.max(bitmap.width,bitmap.height));
    canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(bitmap.width*scale));canvas.height=Math.max(1,Math.round(bitmap.height*scale));
    const context=canvas.getContext('2d');if(!context)throw new Error('Your browser could not prepare this image.');
    context.drawImage(bitmap,0,0,canvas.width,canvas.height);
    const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));
    if(!blob)throw new Error('Your browser could not prepare this image.');
    return {name:file.name,data:new Uint8Array(await blob.arrayBuffer()),type:'png',width:canvas.width,height:canvas.height};
  } catch(error) {
    if(/40 megapixels|Your browser/.test(error.message))throw error;
    throw new Error(`“${file.name}” is not a readable ${extension(file.name).toUpperCase()} image. Try saving it as JPG or PNG.`);
  } finally {bitmap?.close();if(canvas){canvas.width=0;canvas.height=0;}}
}

export async function imageDocument(images,opt={}) {
  if(!images.length||images.length>30)throw new Error('Choose between 1 and 30 images.');
  const {Document,Packer,Paragraph,ImageRun,TextRun,AlignmentType,SectionType,PageOrientation}=await import('docx');
  const sections=images.map(img=>{
    const layout=fitImage(img.width,img.height,opt),caption=opt.imageCaptions!=='no';
    return {properties:{type:SectionType.NEXT_PAGE,page:{size:{width:Math.round(Math.min(...layout.page)*20),height:Math.round(Math.max(...layout.page)*20),orientation:opt.orientation==='landscape'?PageOrientation.LANDSCAPE:PageOrientation.PORTRAIT},margin:{top:840,bottom:840,left:840,right:840}}},children:[
      ...(caption?[new Paragraph({keepNext:true,spacing:{after:160},children:[new TextRun({text:img.name.length>80?img.name.slice(0,77)+'…':img.name,bold:true,size:20,font:'Arial'})]})]:[]),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new ImageRun({data:img.data,type:img.type,transformation:{width:layout.width,height:layout.height},altText:{title:img.name,description:img.name,name:img.name}})]})
    ]};
  });
  return Packer.toBlob(new Document({creator:'Just a File Converter',title:'Image document',styles:{default:{document:{run:{font:'Arial',size:20},paragraph:{spacing:{before:0,after:0}}}}},sections}));
}

export async function imagesToWord(files,opt={},progress=()=>{}) {
  validateFiles(files,IMAGE_WORD_EXTENSIONS);
  if(!files.length||files.length>30)throw new Error('Choose between 1 and 30 images.');
  const images=[];let totalBytes=0;
  for(let i=0;i<files.length;i++) {
    const img=await normalizeImage(files[i]);totalBytes+=img.data.byteLength;
    if(totalBytes>100*1024*1024)throw new Error('The prepared images are too large for one document. Choose fewer or smaller images.');
    images.push(img);progress(10+(i+1)/files.length*70,`Preparing image ${i+1} of ${files.length}…`);
    await new Promise(resolve=>setTimeout(resolve,0));
  }
  progress(85,'Creating your image document…');
  return {name:(files.length===1?fileStem(files[0].name):'image-document')+'.docx',blob:await imageDocument(images,opt),note:'One image per page, in your chosen order. Open this .docx in Word. Images remain pictures; text is not extracted (no OCR). Animated files use one still frame. Large images are resized to 3,000 pixels on their longest side.'};
}
