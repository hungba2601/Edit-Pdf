import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';
import fs from 'fs';

async function createPdf() {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([600, 400]);

  const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const timesFont = await pdfDoc.embedFont(StandardFonts.TimesRoman);

  page.drawText('Hello this is a sample PDF file for editing.', {
    x: 50,
    y: 350,
    size: 20,
    font: helveticaFont,
    color: rgb(0, 0, 0),
  });

  page.drawText('You can click here to edit this line of text.', {
    x: 50,
    y: 300,
    size: 14,
    font: timesFont,
    color: rgb(0.1, 0.2, 0.6),
  });

  page.drawText('This line of text uses Helvetica Bold.', {
    x: 50,
    y: 250,
    size: 16,
    font: await pdfDoc.embedFont(StandardFonts.HelveticaBold),
    color: rgb(0.8, 0.1, 0.1),
  });

  page.drawText('This is a smaller text line at the bottom.', {
    x: 50,
    y: 150,
    size: 10,
    font: helveticaFont,
    color: rgb(0.4, 0.4, 0.4),
  });

  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync('sample.pdf', pdfBytes);
  console.log('Created sample.pdf successfully!');
}

createPdf();
