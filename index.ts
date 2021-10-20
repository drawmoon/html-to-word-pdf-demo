import { join } from 'path';
import { readFileSync, writeFileSync } from 'fs';
import { JSDOM } from 'jsdom';
import { loadImage, createCanvas } from 'canvas';
import * as word from 'html-to-docx';
import * as pdf from 'html-pdf-node';

const basePath = process.cwd().replace('dist', '');
const outputPath = process.cwd();

const HTML_TEMPLATE = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Document</title>
  <style>
    table {
      min-height: 25px;
      line-height: 25px;
      text-align: center;
      border-collapse: collapse;
    }
    table,
    table tr,
    table tr th,
    table tr td {
      border: 1px solid #a8aeb2;
      padding: 5px 10px;
    }
    img {
      max-width: 100%;
      max-height: 100%;
    }
    figure.image {
      display: table;
      clear: both;
      text-align: center;
      margin: 1em auto;
    }
  </style>
</head>
<body>
  {htmlContent}
</body>
</html>`;

const targetWidth = 595;
const targetHeight = 842;

function imgZoom(imgWidth: number, imgHeight: number): { width: number; height: number } {
    let width: number;
    let height: number;

    if (imgWidth > 0 && imgHeight > 0) {
        //原图片宽高比例 大于 指定的宽高比例，这就说明了原图片的宽度必然 > 高度
        if (imgWidth / imgHeight >= targetWidth / targetHeight) {
            if (imgWidth > targetWidth) {
                width = targetWidth;
                // 按原图片的比例进行缩放
                height = (imgHeight * targetWidth) / imgWidth;
            } else {
                // 按照图片的大小进行缩放
                width = imgWidth;
                height = imgHeight;
            }
        } else {
            // 原图片的高度必然 > 宽度
            if (imgHeight > targetHeight) {
                height = targetHeight;
                // 按原图片的比例进行缩放
                width = (imgWidth * targetHeight) / imgHeight;
            } else {
                // 按原图片的大小进行缩放
                width = imgWidth;
                height = imgHeight;
            }
        }
    }

    return { width, height };
}

async function convertImg(html: string): Promise<string> {
    const dom = new JSDOM(html);
    const doc = dom.window.document;

    console.log('Converting image file to base64.');

    for (const img of doc.querySelectorAll('img')) {
        const imgFile = join(basePath, img.src);

        const svg = await loadImage(imgFile);
        console.log(`Width: ${svg.width}, Height: ${svg.height}`);

        if (svg.width > targetWidth || svg.height > targetHeight) {
            console.log('The image is out of bounds.');

            const { width, height } = imgZoom(svg.width, svg.height);

            console.log(`Redrawing image, Width: ${width}, Height: ${height}`);

            const canvas = createCanvas(width, height);
            const context = canvas.getContext('2d');
            context.drawImage(svg, 0, 0, width, height);

            img.src = canvas.toDataURL('image/png');
        } else {
            const imgBase64 = readFileSync(imgFile, { encoding: 'base64' });
            img.src = `data:image/png;base64,${imgBase64}`;
        }
    }

    return doc.body.innerHTML;
}

async function readHtml(): Promise<string> {
    let htmlStr = readFileSync(join(basePath, 'assets', 'content.html'), {
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

// 生成 docx 文件
toWord().then((buff) => {
    writeFileSync(join(outputPath, 'document.docx'), buff);
});

// 生成 pdf 文件
toPdf().then((buff) => {
    writeFileSync(join(outputPath, 'document.pdf'), buff);
})
