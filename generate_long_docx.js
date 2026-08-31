const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Header, Footer, PageNumber, LevelFormat, PageBreak } = require('docx');

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
        run: { size: 36, bold: true, font: "Microsoft YaHei", color: "000000" },
        paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 }
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: "Microsoft YaHei", color: "000000" },
        paragraph: { spacing: { before: 280, after: 180 }, outlineLevel: 1 }
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 24, bold: true, font: "Microsoft YaHei", color: "000000" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      },
      {
        reference: "numbers",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
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
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "西灵呓语", font: "Microsoft YaHei", size: 18, color: "888888" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "第 ", font: "Microsoft YaHei", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Microsoft YaHei", size: 18, color: "888888" }),
            new TextRun({ text: " 页", font: "Microsoft YaHei", size: 18, color: "888888" })
          ]
        })]
      })
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
        children: [new TextRun({ text: "一份关于低强度表达与边缘记忆的长期书写计划", font: "Microsoft YaHei", size: 22, color: "666666" })]
      }),
      new Paragraph({
        spacing: { before: 200, after: 200, line: 400 },
        children: [new TextRun({ text: "在光与雾的交界，听见另一种低语。这里记录虚构与真实的碎片：故事、影像、随想，以及尚未成形的念头。它不是一份明确的宣言，也不是一条被设计好的路线图，而更像是一片被允许缓慢生长的自留地。作者希望它既属于私人记录，也具备某种可被分享的质地。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 200, after: 200, line: 400 },
        children: [new TextRun({ text: "这篇文章本身，就是对这个空间的一次展开说明。它想回答几个朴素的问题：西灵呓语到底是什么？它为什么存在？它目前在做些什么，以及接下来可能走向哪里？这些问题没有标准答案，但书写本身就是一种靠近答案的方式。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("一、命名与起源")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "“西灵呓语”这个名字，最初来自一次对边界的直觉性想象。西，常常意味着远方、黄昏、退向一侧的视线；灵，则保留着某种尚未被完全解释的感知；而呓语，就是那种并不追求被所有人听见的低语。三者放在一起，并不指向某个确切的地理位置，也不是某种现成的文化符号，而更像是一种情感上的坐标。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "它诞生于一个简单却持续的需求：在信息越来越快、表达越来越响的时代里，保留一处允许缓慢、模糊、不完整发生的地方。这个需求并不特别，很多写作者、记录者、影像制作者都可能有过类似的念头。但对这个项目而言，它被具体化为一个持续更新的站点，以及随之而来的若干写作与整理习惯。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "从更个人的层面看，这个名字也与某些长期兴趣有关：对夜间环境、薄雾、旧建筑、边角街道、低分辨率记忆的兴趣。它们未必会全部进入最终作品，但它们持续塑造着这个项目的审美与节奏。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("二、核心理念：低强度表达")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "“低强度表达”是这个项目最核心也最容易被误解的概念。它不等于“随便写写”，也不等于“情绪化”或“碎片化”。它的真正含义是：尽量少地使用那种能够快速抓取注意力、却也在同时消耗注意力的表达方式；尽量多地为读者留下停顿、歧义和再次进入的空间。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "从实践上看，这意味着几个具体选择：第一，篇幅通常不长，但句子往往更慢；第二，避免强烈的结论性语气；第三，允许不确定感留在文本中；第四，不追求“有用”，更在意“可感”。这些选择不是规则，而是倾向。它们偶尔会被打破，但整体上构成一种可识别的语气。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这种倾向也解释了为什么这里不会频繁更新，也不会追逐热点。它不是一份新闻简报，也不是一个观点专栏。它更像一间深夜仍亮着灯的房间，访客随时可以推门进来，但主人并不保证每次都准备好谈话。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("三、内容方向")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "目前，西灵呓语的内容大致落在三个重叠又不完全相同的区域里：叙事、氛围图片、短文。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("1. 叙事")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "叙事不是传统意义上的故事写作。这里的叙事更接近“场景记录”或“情绪结构”：它们有开头、发展和某种松动意义上的结尾，但重点往往不在事件，而在事件发生时的空气密度。它可能是一段夜间行车见闻，一栋老楼里的脚步声，或者一封没有寄出去的信的片段。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这些叙事片段通常较短，但它们彼此之间可能存在隐线。随着时间推移，这些隐线可能会连成更大的结构，也可能始终保持分离。作者并不预先决定答案。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("2. 氛围图片")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "图片在这里不是插图，而是与文字平行的表达系统。它们不总是高分辨率、高饱和、高信息量的“作品”，而更接近某种观察痕迹：雾里的路灯、半开的门、水面上的反光、被风吹歪的招牌。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这些图片通常会尽量保留拍摄时的环境信息，例如光线、颗粒感、曝光失误带来的偏色，以及构图上的犹豫。它们不追求“完美”，而是追求“在场”。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("3. 短文")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "短文是介于笔记与文章之间的文体。它们比笔记更有结构，比文章更短、更私人。话题可以是一本书的某一页、一首歌的某个和弦、一种颜色在傍晚时分的变化，或者某种难以命名的情绪。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这些短文不会定期发布，也不会按照主题分类整理。它们会在合适的时机出现，像落在桌上的一片叶子。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("四、目前的作品与状态")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "到目前为止，西灵呓语仍处于早期阶段。网站已经建立，但内容以占位与试写为主。已有的文字片段主要围绕“观察”展开，图片则集中在少数几次夜间与清晨的拍摄。这个阶段的重要任务不是“完成”，而是“建立语气”：让未来的自己与读者知道，这里的声音大概会是什么样子。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "从结构上看，目前的页面分为几个固定区域：首页、关于、作品、呓语。首页负责建立第一印象；关于部分解释项目理念与长期方向；作品区将逐步替换为真实内容；呓语区则留给最短的思考、未完成句子，或一句今晚想留住的话。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这种结构有意保持简单。它不是为了展示复杂性，而是为了不给内容本身造成负担。随着内容增加，结构可能会调整，但原则依然是：形式退后，让内容呼吸。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("五、写作习惯与工作方式")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "西灵呓语的写作并不是“坐在书桌前开始写”。更多时候，它发生在散步、通勤、深夜刷到某张旧照片、或者听到一段与环境不符的音乐时。这些时刻会产生一些零散材料：一段语音备忘录、几张手机照片、几个写在纸上的词，或者一个没有标题的文档。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这些材料不会立刻被整理成文章。它们会先在某个临时空间里共存一段时间，也许是几天，也许是几个月。等到某个安静的下午或深夜，作者会重新打开它们，看看哪些仍然有效，哪些已经失去颜色，哪些开始彼此说话。那个重新打开的瞬间，才是真正的写作开始。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这也意味着，发布节奏完全由内容决定，而不是由时间表决定。这里没有“每周更一篇”的承诺，也没有“月末总结”的压力。唯一的要求是：每当某样东西被放出来时，它应该是经过时间检验的。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("六、技术选择与限制")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "作为一个主要面向文字与图片的项目，技术选择上刻意保持克制。目前站点使用静态页面构建，没有数据库，没有评论系统，没有订阅功能。这样的选择并非出于对技术的排斥，而是出于对“可控性”的偏好。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "静态站点意味着维护成本低，也意味着内容不会因为某个插件更新而崩溃。它像一个安静的仓库，只存放已经整理好的东西。而动态的讨论、反馈、协作，则留给更合适的外部空间，例如邮件、社交平台，或更专门的论坛。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这种克制也延伸到文档生成。当前Word文档的生成遵循同样原则：清晰、标准、不过度装饰。在未来的内容更丰富之后，排版风格可能会随之演化，但标准化的原则应该会保留。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("七、后续计划")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "西灵呓语接下来的工作，会围绕几个仍在酝酿中的项目展开。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("“雾中小径”摄影系列")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这是一组以清晨与夜间雾气为背景的街景与自然片段。拍摄地点尚未完全确定，但目前集中在城市边缘带、河岸、以及一些正在被遗忘的旧街区。这个系列的目标不是“记录一座城市”，而是记录城市在特定光线与湿度条件下，那种介于清醒与梦境之间的状态。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("“来信”读者互动栏目")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这个栏目不是传统意义上的问答或投稿，而更像一场持续的书信往还。读者可以寄来一句话、一段记忆、一张照片、一段音频，或者任何他们认为适合放在这里的材料。作者会进行筛选、整理，并在之后的页面中以某种形式回应。这个栏目的核心不是“互动量”，而是“信任感”。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("夜间写作实验")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "不定期的夜间写作实验，会在特定时间段内进行。规则很简单：从深夜开始写，到清晨结束，中间不修改、不回头。写出来的东西可能完整，可能断裂，可能充满重复与自我矛盾。但它们至少保留了某个夜晚真实的呼吸节奏。这些实验文本不会被单独包装成“作品”，而是作为呓语的一部分自然出现。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("八、给访客的话")] }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "如果你偶然来到这里，说明你和我一样，仍然对某种慢下来、静下来、保留不确定性的方式有兴趣。你不需要点赞，不需要订阅，不需要留下任何痕迹。你只需要读，然后离开；或者留下，然后离开。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "这里的内容会持续生长，也会持续修正。如果你发现某段文字让你停下来，那就说明它完成了最基本的任务。", font: "Microsoft YaHei", size: 24 })]
      }),
      new Paragraph({
        spacing: { before: 120, after: 200, line: 400 },
        children: [new TextRun({ text: "最后，感谢你读到这里。", font: "Microsoft YaHei", size: 24 })]
      }),

      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 400 },
        children: [new TextRun({ text: "西灵呓语  写于某个没有特别之处的夜晚", font: "Microsoft YaHei", size: 22, color: "888888" })]
      })
    ]
  }]
});

const outPath = '/run/csi/mount-root/nas/4079184d856ecc166ed19d4887083405/workspaces/default/xilingyiyu/西灵呓语.docx';
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log('written', outPath);
});
