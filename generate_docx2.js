const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat } = require('docx');

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Microsoft YaHei", size: 24 }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 36, bold: true, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 }
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 280, after: 180 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [new TextRun("西灵呓语")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 360 },
        children: [new TextRun({ text: "一份简单的介绍", font: "Microsoft YaHei", size: 22, color: "666666" })]
      }),
      new Paragraph({
        spacing: { before: 200, after: 200, line: 400 },
        children: [new TextRun("西灵呓语是一个虚构与纪实并置的空间。核心理念是“低强度表达”：不追求喧嚣，而留意那些容易被忽略的声音。")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("方向")]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun("目前更偏重叙事、氛围图片与短文。后续可能加入声音片段、极短影像，以及合作栏目。")]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("计划")]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun("例如：“雾中小径”摄影系列、“来信”读者互动栏目，以及不定期的夜间写作实验。")]
      })
    ]
  }]
});

const outPath = '/run/csi/mount-root/nas/4079184d856ecc166ed19d4887083405/workspaces/default/xilingyiyu/西灵呓语.docx';
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log('written', outPath);
});
