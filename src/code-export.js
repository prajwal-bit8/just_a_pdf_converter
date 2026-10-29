import {prepareCode, missingPdfGlyph} from './code.js';

const creator = 'Just a File Converter';
const fontName = 'DejaVuSansMono.ttf';
let fontPromise;
async function getFont() {
  if (!fontPromise) fontPromise = fetch('/fonts/'+fontName).then(async response=>{
    if (!response.ok) throw new Error('The code font could not load. Check your connection and try again.');
    return new Uint8Array(await response.arrayBuffer());
  }).catch(error=>{fontPromise=null;throw error;});
  return fontPromise;
}

export async function codeToPdf(files, options={}, progress=()=>{}, fontBytes) {
  const prepared = await prepareCode(files,options), {opt,documents,name}=prepared;
  progress(20,'Formatting your source code…');
  const [{default:pdfMake},font] = await Promise.all([import('pdfmake/build/pdfmake.js'),fontBytes || getFont()]);
  const missing = missingPdfGlyph(font, opt.title + documents.map(d=>d.name+d.language+d.text).join(''));
  if (missing) throw new Error(`The PDF font cannot display “${missing}”. Choose Code to Word and a font that supports this text in Word, or replace the character and try again.`);
  let binary='';
  for (let i=0;i<font.length;i+=8192) binary+=String.fromCharCode(...font.subarray(i,i+8192));
  const content=[];
  documents.forEach((doc,index)=>{
    content.push({text:opt.title || 'Source code',fontSize:16,color:'#493582',margin:[0,0,0,8],pageBreak:index?'before':undefined});
    content.push({text:doc.name,fontSize:11,margin:[0,0,0,4]});
    content.push({text:`${doc.language} · ${doc.lineCount} lines`,fontSize:9,color:'#666666',margin:[0,0,0,14]});
    const gutter=opt.lineNumbers?(doc.digits+3)*opt.fontSize*.62:0;
    content.push({table:{widths:opt.lineNumbers?[gutter,'*']:['*'],body:doc.rows.map(row=>{
      const code={text:row.text || ' ',noWrap:true,preserveLeadingSpaces:true,preserveTrailingSpaces:true};
      return opt.lineNumbers?[{text:row.number,color:'#777777',alignment:'right',margin:[0,0,10,0]},code]:[code];
    }),dontBreakRows:true},layout:{hLineWidth:()=>0,vLineWidth:()=>0,paddingLeft:()=>0,paddingRight:()=>0,paddingTop:()=>0,paddingBottom:()=>2}});
  });
  const definition={pageSize:opt.paper,pageOrientation:opt.landscape?'landscape':'portrait',pageMargins:[42,42,42,42],
    defaultStyle:{font:'CodeMono',fontSize:opt.fontSize},content,
    footer:(page,total)=>({text:`${page} / ${total}`,alignment:'center',fontSize:8,margin:[0,14,0,0],color:'#777777'}),
    info:{title:opt.title || name,creator}};
  progress(60,'Creating your PDF…');
  const blob=await new Promise((resolve,reject)=>{
    try {
      const stream=pdfMake.createPdf(definition,null,{CodeMono:{normal:fontName}},{[fontName]:btoa(binary)}).getStream(),chunks=[];
      stream.on('data',chunk=>chunks.push(new Uint8Array(chunk)));stream.on('error',reject);
      stream.on('end',()=>resolve(new Blob(chunks,{type:'application/pdf'})));stream.end();
    } catch(error){reject(error);}
  });
  return {name:name+'.pdf',blob,note:'Source code formatted with indentation preserved and long lines wrapped. Each source file starts on a new page. Code is not executed.'};
}

export async function codeToWord(files, options={}, progress=()=>{}) {
  const {opt,documents,name} = await prepareCode(files,options);
  const {Document,Packer,Paragraph,TextRun,Footer,PageNumber,AlignmentType,SectionType,LineRuleType,PageOrientation}=await import('docx');
  progress(30,'Creating your editable code document…');
  const output=new Document({creator,title:opt.title || name,
    styles:{default:{document:{run:{font:'Courier New',size:opt.fontSize*2},paragraph:{spacing:{before:0,after:0}}}}},
    sections:documents.map(doc=>({properties:{type:SectionType.NEXT_PAGE,page:{size:{width:Math.round(Math.min(...opt.dimensions)*20),height:Math.round(Math.max(...opt.dimensions)*20),orientation:opt.landscape?PageOrientation.LANDSCAPE:PageOrientation.PORTRAIT},margin:{top:840,bottom:840,left:840,right:840}}},
      footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({children:[PageNumber.CURRENT,' / ',PageNumber.TOTAL_PAGES],size:16,color:'777777'})]})]})},
      children:[
        new Paragraph({keepNext:true,spacing:{after:160},children:[new TextRun({text:opt.title || 'Source code',font:'Arial',size:32,bold:true,color:'493582'})]}),
        new Paragraph({keepNext:true,spacing:{after:80},children:[new TextRun({text:doc.name,font:'Arial',size:22,bold:true})]}),
        new Paragraph({keepNext:true,spacing:{after:240},children:[new TextRun({text:`${doc.language} · ${doc.lineCount} lines`,font:'Arial',size:18,color:'666666'})]}),
        ...doc.rows.map(row=>new Paragraph({widowControl:false,spacing:{before:0,after:0,line:Math.round(opt.fontSize*1.35*20),lineRule:LineRuleType.EXACT},children:[
          ...(opt.lineNumbers?[new TextRun({text:row.number.padStart(doc.digits)+'   ',color:'777777'})]:[]),
          new TextRun({text:row.text || ' '})
        ]}))
      ]}))});
  return {name:name+'.docx',blob:await Packer.toBlob(output),note:'Editable source code in Word, with indentation preserved and long lines wrapped. Each source file starts on a new page. For other scripts, choose a font in Word that supports those characters. Code is not executed.'};
}
