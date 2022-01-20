import { join } from 'path';
import { readFileSync, writeFileSync } from 'fs';
import { JSDOM } from 'jsdom';
import { loadImage, createCanvas } from 'canvas';
import * as word from 'html-to-docx';
import * as pdf from 'html-pdf-node';

const basePath = join(process.cwd().replace('dist', ''), 'assets');
const outputPath = process.cwd();

const HTML_TEMPLATE = `
<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <title>Document</title>
  <style>
    body {
      -webkit-print-color-adjust: exact;
    }
    p {
      display: block;
      margin-left: 2em;
    }
    figure.table {
      display: table;
    }
    table {
      display: table;
      text-indent: initial;
      border-collapse: collapse;
      border-spacing: 0;
      width: 100%;
      height: 100%;
      border: 1px double #b3b3b3;
    }
    table tr {
      display: table-row;
      vertical-align: inherit;
      border-color: inherit;
    }
    table td,
    table th {
      min-width: 2em;
      padding: .4em;
      border: 1px solid #bfbfbf;
    }
    figure.image {
        display: table;
    }
    img {
      display: table-caption;
      margin: 0 auto;
      max-width: 100%;
      min-width: 100%;
    }
    figcaption {
      display: table-caption;
      word-break: break-word;
      text-align: center;
      color: #333;
      background-color: #f7f7f7;
      padding: .6em;
      font-size: .75em;
      outline-offset: -1px;
    }
    .table figcaption {
      caption-side: top;
    }
    .image figcaption {
      caption-side: bottom;
    }
    .todo-list {
      list-style: none;
      margin-left: -1.5em;
    }
  </style>
</head>
<body>
  {htmlContent}
</body>
</html>`;

const expectImageWidth = 595;
const expectImageHeight = 842;

async function resize(file: string): Promise<string> {
    const svg = await loadImage(file);

    let width = svg.width;
    let height = svg.height;

    if (width > expectImageWidth) {
        height = (height * expectImageWidth) / width;
        width = expectImageWidth;
    } else if (height > expectImageHeight) {
        width = (width * expectImageHeight) / height;
        height = expectImageHeight;
    }

    console.log(`Redrawing image, Width: ${width}, Height: ${height}`);

    const canvas = createCanvas(width, height);
    const context = canvas.getContext('2d');
    context.drawImage(svg, 0, 0, width, height);
    return canvas.toDataURL('image/png');
}

async function convertImg(html: string): Promise<string> {
    const dom = new JSDOM(html);
    const doc = dom.window.document;

    console.log('Converting image file to base64.');

    for (const img of doc.querySelectorAll('img')) {
        const imgFile = join(basePath, img.src);

        img.src = await resize(imgFile);
    }

    return doc.body.innerHTML;
}

async function readHtml(): Promise<string> {
    let htmlStr = readFileSync(join(basePath, 'index.html'), {
        encoding: 'utf-8',
    });

    htmlStr = await convertImg(htmlStr);

    return htmlStr;
}

async function toPdf(): Promise<Buffer> {
    const htmlContent = await readHtml();
    const html = HTML_TEMPLATE.replace('{htmlContent}', htmlContent);

    return await pdf.generatePdf(
        { content: html },
        { format: 'A4', margin: { top: 40, right: 20, bottom: 40, left: 20 } },
    );
}

async function toWord(): Promise<Buffer> {
    const htmlContent = await readHtml();
    const html = HTML_TEMPLATE.replace('{htmlContent}', htmlContent);

    return await word(html, null, {
        table: { row: { cantSplit: true } },
    });
}

// 保存 docx 文件
toWord().then((buff) => {
    writeFileSync(join(outputPath, 'dist/document.docx'), buff);
});

// 保存 pdf 文件
toPdf().then((buff) => {
    writeFileSync(join(outputPath, 'dist/document.pdf'), buff);
});
