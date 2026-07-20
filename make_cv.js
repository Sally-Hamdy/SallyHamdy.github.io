const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, ImageRun,
  LevelFormat, convertInchesToTwip, TabStopType
} = require('docx');

const ACCENT = "7B2CBF";   // deep violet
const DARK = "1A1A2E";
const GRAY = "555566";
const LIGHT = "F3EEFA";

const noBorders = { top:{style:BorderStyle.NONE,size:0,color:"FFFFFF"}, bottom:{style:BorderStyle.NONE,size:0,color:"FFFFFF"}, left:{style:BorderStyle.NONE,size:0,color:"FFFFFF"}, right:{style:BorderStyle.NONE,size:0,color:"FFFFFF"} };

function sectionTitle(text){
  return new Paragraph({
    spacing:{before:260,after:120},
    border:{bottom:{style:BorderStyle.SINGLE,size:8,color:ACCENT,space:2}},
    children:[new TextRun({text:text.toUpperCase(),bold:true,size:24,color:ACCENT,font:"Calibri",characterSpacing:20})]
  });
}
function job(title, org, when, desc){
  const out=[new Paragraph({
    spacing:{before:140,after:20},
    tabStops:[{type:TabStopType.RIGHT,position:convertInchesToTwip(7.0)}],
    children:[
      new TextRun({text:title,bold:true,size:22,color:DARK,font:"Calibri"}),
      new TextRun({text:" — "+org,size:22,color:DARK,font:"Calibri"}),
      new TextRun({text:"\t"+when,size:20,color:ACCENT,bold:true,font:"Calibri"})
    ]})];
  if(desc) out.push(new Paragraph({spacing:{after:40},children:[new TextRun({text:desc,size:20,color:GRAY,font:"Calibri"})]}));
  return out;
}
function bullet(text){
  return new Paragraph({numbering:{reference:"b",level:0},spacing:{after:40},children:[new TextRun({text,size:20,color:GRAY,font:"Calibri"})]});
}

const photo = fs.readFileSync('sally.jpeg');

const headerTable = new Table({
  width:{size:convertInchesToTwip(7.0),type:WidthType.DXA},
  columnWidths:[convertInchesToTwip(1.4),convertInchesToTwip(5.6)],
  borders:noBorders,
  rows:[new TableRow({children:[
    new TableCell({
      width:{size:convertInchesToTwip(1.4),type:WidthType.DXA},
      borders:noBorders, margins:{top:60,bottom:60,left:0,right:120},
      children:[new Paragraph({children:[new ImageRun({data:photo,type:"jpg",transformation:{width:95,height:122}})]})]
    }),
    new TableCell({
      width:{size:convertInchesToTwip(5.6),type:WidthType.DXA},
      borders:noBorders, margins:{top:60,bottom:60,left:120,right:0},
      children:[
        new Paragraph({spacing:{after:40},children:[new TextRun({text:"SALSABEEL HAMDY",bold:true,size:52,color:DARK,font:"Calibri"})]}),
        new Paragraph({spacing:{after:100},children:[new TextRun({text:"Social Media Coordinator  ·  Content Creator  ·  AI Video Creator",size:24,color:ACCENT,bold:true,font:"Calibri"})]}),
        new Paragraph({spacing:{after:30},children:[new TextRun({text:"Sharjah, UAE   |   salsabeel.h.mohamed60@gmail.com   |   +971 50 229 5504",size:19,color:GRAY,font:"Calibri"})]}),
        new Paragraph({children:[new TextRun({text:"Portfolio: sally-hamdy.github.io/SallyHamdy.github.io",size:19,color:GRAY,font:"Calibri"})]})
      ]
    })
  ]})]
});

const doc = new Document({
  numbering:{config:[{reference:"b",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:280,hanging:160}}}}]}]},
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:720,bottom:720,left:720,right:720}}},
    children:[
      headerTable,

      sectionTitle("Profile"),
      new Paragraph({children:[new TextRun({text:"Social Media Coordinator and Content Creator with 5+ years of experience growing brands across very different industries — bridal fashion, e-commerce, industrial packaging, and technology. I build creative strategies, produce engaging multimedia content, and create modern AI-generated videos that give brands a cutting-edge presence. Managed the channels of 7 brands, from strategy and content calendars to community engagement and analytics.",size:20,color:GRAY,font:"Calibri"})]}),

      sectionTitle("Experience"),
      ...job("Social Media Manager","LinSync Technologies","2025 – Present","Managing all media channels for a technology company — content strategy, publishing, brand voice, and AI-powered video content."),
      ...job("Social Media Manager","Dupack & Empack (Packaging Factories, UAE)","2024 – 2025","Managed the social channels of two packaging factories, building their B2B presence and producing AI-generated promotional videos of their products and facilities."),
      ...job("Social Media Coordinator (Freelance)","Bookbee","2024 – Present","Brand identity — logos, color schemes, typography — plus graphics for web, social media, and digital ads."),
      ...job("Content Creator (Freelance)","DropShippingScout","2023 – 2024","Developed visual concepts aligned with project goals; supervised the design process and managed timelines and budgets."),
      ...job("Social Media Coordinator","Tatweer Company","2022 – 2023","Responsible for all social media channels, including LinkedIn and Facebook."),
      ...job("Social Media Coordinator","TiaraBridal","2021 – 2023","Ran all social channels for a bridal fashion brand, growing engagement with on-trend content."),

      sectionTitle("Skills"),
      bullet("Social media strategy & management (Instagram, Facebook, LinkedIn, TikTok)"),
      bullet("Content creation — graphics, copywriting, multimedia campaigns"),
      bullet("AI video production & editing with the latest AI tools"),
      bullet("Brand identity design — logos, color schemes, typography"),
      bullet("Web design — clean, modern, responsive"),
      bullet("Analytics, trends & community engagement"),

      sectionTitle("Education"),
      ...job("Bachelor of Science","Cairo University","2018 – 2022","Foundation in digital communication, media production, and data analysis, with focus on content strategy, visual storytelling, and audience engagement."),

      sectionTitle("Certifications"),
      new Paragraph({spacing:{before:100,after:60},children:[new TextRun({text:"Meta Social Media Marketing Professional Certificate — Coursera (in progress)",bold:true,size:21,color:DARK,font:"Calibri"})]}),
      bullet("Social Media Management — Completed"),
      bullet("Fundamentals of Social Media Advertising · Advertising with Meta · Measure & Optimize Social Media Marketing Campaigns — In progress")
    ]
  }]
});

Packer.toBuffer(doc).then(b=>{fs.writeFileSync("Salsabeel-Hamdy-CV.docx",b);console.log("done",b.length)});
