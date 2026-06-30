import { PDFDocument, rgb, StandardFonts, PDFName } from 'pdf-lib';

export class PdfSaver {
  static hexToRgb(hex) {
    const r = parseInt(hex.slice(1,3),16)/255;
    const g = parseInt(hex.slice(3,5),16)/255;
    const b = parseInt(hex.slice(5,7),16)/255;
    return rgb(r,g,b);
  }

  static async save(originalFile, annotations, drawingDataUrls, scale, originalTextEdits = {}) {
    const buf = await originalFile.arrayBuffer();
    const pdfDoc = await PDFDocument.load(buf);
    const pages = pdfDoc.getPages();

    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      const { width: pw, height: ph } = page.getSize();
      const pageAnns = annotations[i] || [];
      console.log(`[PdfSaver Debug] Page ${i} - pageAnns:`, JSON.parse(JSON.stringify(pageAnns)));
      const drawDataUrl = drawingDataUrls[i];
      const textEdits = originalTextEdits[i] || [];

      // === Remove original annotations (like digital signature stamps) that overlap with eraser paths ===
      const annots = page.node.Annots();
      if (annots) {
        const getVal = (obj) => {
          if (!obj) return 0;
          if (typeof obj.value === 'number') return obj.value;
          if (typeof obj.numberValue === 'function') return obj.numberValue();
          if (typeof obj === 'number') return obj;
          return Number(obj);
        };

        for (let j = annots.size() - 1; j >= 0; j--) {
          const annot = annots.lookup(j);
          if (!annot) continue;
          
          const rect = annot.get(PDFName.of('Rect'));
          if (rect && rect.size() === 4) {
            const rx1 = getVal(rect.lookup(0));
            const ry1 = getVal(rect.lookup(1));
            const rx2 = getVal(rect.lookup(2));
            const ry2 = getVal(rect.lookup(3));
            
            // Check if this annotation overlaps with any of our eraser paths
            for (const ann of pageAnns) {
              if (ann.type === 'drawing' && ann.path && ann.path.tool === 'eraser') {
                const p = ann.path;
                const ex1 = Math.min(p.startX, p.endX) / scale;
                const eyCanvas1 = Math.min(p.startY, p.endY);
                const ew = Math.abs(p.endX - p.startX) / scale;
                const eh = Math.abs(p.endY - p.startY) / scale;
                const ey1 = ph - (eyCanvas1 / scale) - eh;
                
                const ex2 = ex1 + ew;
                const ey2 = ey1 + eh;
                
                // Check intersection of bounding boxes
                const intersects = !(rx1 > ex2 || rx2 < ex1 || ry1 > ey2 || ry2 < ey1);
                if (intersects) {
                  // Remove this annotation from the array
                  annots.remove(j);
                  break;
                }
              }
            }
          }
        }
      }

      // === Cover original text edits with background-colored rectangles ===
      for (const edit of textEdits) {
        const item = edit.origItem;
        const x = item.x / scale;
        const rectY = ph - (item.y / scale) - (item.height / scale); // Bottom of the box
        const w = item.width / scale;
        const h = item.height / scale;

        const bg = item.bgRgb || { r: 1, g: 1, b: 1 };

        // Draw background-colored rectangle covering the entire text block
        // Use generous padding scaled proportionally to fully cover anti-aliased edges
        const padX = Math.max(3, (item.fontSize || 12) * 0.15) / scale;
        const padY = Math.max(3, (item.fontSize || 12) * 0.2) / scale;
        page.drawRectangle({
          x: x - padX,
          y: rectY - padY,
          width: w + padX * 2 + 1,
          height: h + padY * 2 + 1,
          color: rgb(bg.r, bg.g, bg.b),
          borderWidth: 0,
        });
      }

      // Draw native whiteout rectangles for all eraser paths to cover the text with solid vectors,
      // which prevents the underlying black text from leaking through the PNG's semi-transparent anti-aliased edges.
      for (const ann of pageAnns) {
        if (ann.type === 'drawing' && ann.path && ann.path.tool === 'eraser') {
          const p = ann.path;
          const x = Math.min(p.startX, p.endX) / scale;
          const yCanvas = Math.min(p.startY, p.endY);
          const w = Math.abs(p.endX - p.startX) / scale;
          const h = Math.abs(p.endY - p.startY) / scale;
          const y = ph - (yCanvas / scale) - h;

          page.drawRectangle({
            x: x,
            y: y,
            width: w,
            height: h,
            color: rgb(1, 1, 1),
            borderWidth: 0,
          });
        }
      }

      // Embed drawings and text layer (merged into a transparent PNG from browser canvas)
      if (drawDataUrl) {
        try {
          const resp = await fetch(drawDataUrl);
          const imgBytes = await resp.arrayBuffer();
          const img = await pdfDoc.embedPng(new Uint8Array(imgBytes));
          page.drawImage(img, { x: 0, y: 0, width: pw, height: ph });
        } catch(e) { console.warn('Draw embed err', e); }
      }

      // Embed new image annotations only (text annotations are already inside drawDataUrl)
      for (const ann of pageAnns) {
        if (ann.type === 'image' && ann.dataUrl) {
          try {
            const resp = await fetch(ann.dataUrl);
            const imgBytes = new Uint8Array(await resp.arrayBuffer());
            let img;
            if (ann.dataUrl.includes('image/png')) {
              img = await pdfDoc.embedPng(imgBytes);
            } else {
              img = await pdfDoc.embedJpg(imgBytes);
            }
            const x = ann.x / scale;
            const y = ph - (ann.y / scale) - (ann.h / scale);
            page.drawImage(img, { x, y, width: ann.w / scale, height: ann.h / scale });
          } catch(e) { console.warn('Img embed err', e); }
        }
      }
    }

    const bytes = await pdfDoc.save();
    const blob = new Blob([bytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = originalFile.name.replace('.pdf', '_edited.pdf');
    a.click();
    URL.revokeObjectURL(url);
  }
}
