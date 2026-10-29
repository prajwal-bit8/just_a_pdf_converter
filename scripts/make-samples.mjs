import {PDFDocument,StandardFonts,rgb} from 'pdf-lib';
import {Document,Packer,Paragraph,TextRun,HeadingLevel,Table,TableRow,TableCell,WidthType} from 'docx';
import {writeFile,mkdir} from 'node:fs/promises';
await mkdir('dist/samples',{recursive:true});
async function makePDF(name,title,pages){
 const d=await PDFDocument.create(),f=await d.embedFont(StandardFonts.Helvetica),bold=await d.embedFont(StandardFonts.HelveticaBold);
 d.setTitle(title);d.setCreator('Just a File Converter');
 pages.forEach((lines,i)=>{const p=d.addPage([595,842]);p.drawRectangle({x:0,y:783,width:595,height:59,color:rgb(.45,.32,.85)});p.drawText('JUST A FILE CONVERTER / SAMPLE DOCUMENT',{x:42,y:808,size:10,font:bold,color:rgb(1,1,1)});p.drawText(title,{x:42,y:725,size:25,font:bold,color:rgb(.17,.15,.23)});p.drawText(`Page ${i+1} of ${pages.length}`,{x:42,y:698,size:10,font:f,color:rgb(.55,.51,.6)});let y=645;for(const line of lines){p.drawText(line,{x:42,y,size:12,font:f,color:rgb(.35,.31,.4)});y-=29;}p.drawLine({start:{x:42,y:80},end:{x:553,y:80},thickness:1,color:rgb(.9,.87,.94)});p.drawText('A sample file for trying your everyday PDF tools.',{x:42,y:59,size:9,font:f,color:rgb(.6,.55,.65)});});
 await writeFile('dist/samples/'+name,await d.save());
}
await makePDF('project-brief.pdf','Make room for good ideas.',[
 ['PROJECT BRIEF','A simple plan for a small project.','Goal: bring useful ideas to life, one page at a time.','1. Gather the things that inspire you.','2. Choose one problem worth solving.','3. Make a first version and put it to work.'],
 ['THE NEXT CHAPTER','This is the second page of our project brief.','Use Organize PDF to move this page to the beginning.','Rotate a page, remove one, or bring it back.','Your original document stays exactly as it was.'],
 ['READY TO SHARE','This is the final page of the sample document.','Try splitting out pages 1 and 3.','Or convert this file to Word to edit its text.','Small tasks. Simply done.']
]);
await makePDF('meeting-notes.pdf','Notes from a good meeting.',[['MEETING NOTES','We talked. We planned. We wrote things down.','Next step: make something useful.','Owner: the person with the idea.','Deadline: one good step at a time.']]);
const doc=new Document({creator:'Just a File Converter',sections:[{children:[new Paragraph({text:'A little plan for the weekend.',heading:HeadingLevel.TITLE}),new Paragraph({text:'A sample Word document',heading:HeadingLevel.HEADING_2}),new Paragraph({children:[new TextRun('This is a real .docx file, made for trying '),new TextRun({text:'Word to PDF',bold:true}),new TextRun(' in Just a File Converter.')]}),new Paragraph('A good document does not need to be complicated. A clear heading, a few useful words, and a plan are enough.'),new Paragraph({text:'Three things to make time for',heading:HeadingLevel.HEADING_2}),new Paragraph({text:'Learn one new thing.',bullet:{level:0}}),new Paragraph({text:'Build something small.',bullet:{level:0}}),new Paragraph({text:'Spend some time outside.',bullet:{level:0}}),new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[new TableRow({children:['Day','One good thing'].map(text=>new TableCell({children:[new Paragraph(text)]}))}),new TableRow({children:['Saturday','Make a first version'].map(text=>new TableCell({children:[new Paragraph(text)]}))}),new TableRow({children:['Sunday','Review and improve it'].map(text=>new TableCell({children:[new Paragraph(text)]}))})]}),new Paragraph('Your next big idea starts with a small step.')]}]});
await writeFile('dist/samples/weekend-plans.docx',await Packer.toBuffer(doc));
