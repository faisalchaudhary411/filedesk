// FileDesk Shared Tools - All PDF processing logic
// This file is loaded by every tool page

(function(){
"use strict";

if (!window.FileDeskTools) window.FileDeskTools = {};

var U = window.FileDeskUtils;

/* ============== HELPERS ============== */

function readAsArrayBuffer(file){
  return U.readAsArrayBuffer(file);
}

function baseName(name){
  return U.baseName(name);
}

/* ============== RENDER PAGES TO JPEGs (shared by compress, protect, unlock) ============== */

function renderPagesToJpegs(pdfjsDoc, scale, quality, progress, progressBase, progressSpan){
  var count = pdfjsDoc.numPages;
  var pages = [];
  var chain = Promise.resolve();
  for (var i=1;i<=count;i++){
    (function(pageNum){
      chain = chain.then(function(){
        progress(progressBase + (progressSpan*pageNum/count), "Rendering page " + pageNum + " of " + count + "…");
        return pdfjsDoc.getPage(pageNum).then(function(page){
          var viewport = page.getViewport({scale: scale});
          var canvas = document.createElement("canvas");
          canvas.width = viewport.width; canvas.height = viewport.height;
          var ctx = canvas.getContext("2d");
          return page.render({canvasContext: ctx, viewport: viewport}).promise.then(function(){
            pages.push({
              dataUrl: canvas.toDataURL("image/jpeg", quality),
              width: viewport.width,
              height: viewport.height
            });
          });
        });
      });
    })(i);
  }
  return chain.then(function(){ return pages; });
}

/* ============== BUILD DOCX ============== */

function buildDocx(paragraphs){
  function escXml(s){
    return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  }
  var bodyXml = paragraphs.map(function(p){
    if (p === "") return "<w:p/>";
    return "<w:p><w:r><w:t xml:space=\"preserve\">" + escXml(p) + "</w:t></w:r></w:p>";
  }).join("");

  var documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body>' + bodyXml +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr>' +
    '</w:body></w:document>';

  var contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';

  var rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  var docRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

  var zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypes);
  zip.folder("_rels").file(".rels", rootRels);
  var word = zip.folder("word");
  word.file("document.xml", documentXml);
  word.folder("_rels").file("document.xml.rels", docRels);

  return zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
}

/* ============== BUILD PDF FROM TEXT ============== */

function buildPdfFromText(paragraphs){
  var PDFDocument = PDFLib.PDFDocument;
  var StandardFonts = PDFLib.StandardFonts;
  return PDFDocument.create().then(function(doc){
    return doc.embedFont(StandardFonts.Helvetica).then(function(font){
      var pageWidth = 595.28, pageHeight = 841.89;
      var margin = 56, fontSize = 11, lineHeight = 16;
      var maxWidth = pageWidth - margin*2;
      var page = doc.addPage([pageWidth, pageHeight]);
      var y = pageHeight - margin;

      function newPage(){
        page = doc.addPage([pageWidth, pageHeight]);
        y = pageHeight - margin;
      }
      function wrapLine(text){
        var words = text.split(" ");
        var lines = []; var cur = "";
        words.forEach(function(w){
          var test = cur ? cur + " " + w : w;
          var width = font.widthOfTextAtSize(test, fontSize);
          if (width > maxWidth && cur){ lines.push(cur); cur = w; } else { cur = test; }
        });
        if (cur) lines.push(cur);
        return lines.length ? lines : [""];
      }

      paragraphs.forEach(function(para){
        var lines = para.trim() === "" ? [""] : wrapLine(para);
        lines.forEach(function(line){
          if (y < margin){ newPage(); }
          page.drawText(line, {x: margin, y: y, size: fontSize, font: font});
          y -= lineHeight;
        });
        y -= lineHeight * 0.3;
      });

      return doc.save();
    });
  });
}

/* ============== TOOL: MERGE PDF ============== */

window.FileDeskTools["merge-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  return PDFDocument.create().then(function(merged){
    var chain = Promise.resolve();
    files.forEach(function(f, i){
      chain = chain.then(function(){
        progress(10 + (70*i/files.length), "Merging " + f.name + "…");
        return readAsArrayBuffer(f).then(function(buf){
          return PDFDocument.load(buf);
        }).then(function(src){
          return merged.copyPages(src, src.getPageIndices());
        }).then(function(pages){
          pages.forEach(function(p){ merged.addPage(p); });
        });
      });
    });
    return chain.then(function(){
      progress(90, "Saving…");
      return merged.save();
    }).then(function(bytes){
      return [{ name:"merged.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
    });
  });
};

/* ============== TOOL: SPLIT PDF ============== */

window.FileDeskTools["split-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var file = files[0];
  return readAsArrayBuffer(file).then(function(buf){
    return PDFDocument.load(buf);
  }).then(function(src){
    var count = src.getPageCount();
    var zip = new JSZip();
    var chain = Promise.resolve();
    for (var i=0;i<count;i++){
      (function(i){
        chain = chain.then(function(){
          progress(10 + (80*i/count), "Extracting page " + (i+1) + " of " + count + "…");
          return PDFDocument.create().then(function(doc){
            return doc.copyPages(src, [i]).then(function(pages){
              doc.addPage(pages[0]);
              return doc.save();
            }).then(function(bytes){
              zip.file(baseName(file.name) + "-page-" + (i+1) + ".pdf", bytes);
            });
          });
        });
      })(i);
    }
    return chain.then(function(){
      progress(95, "Zipping…");
      return zip.generateAsync({type:"blob"});
    }).then(function(blob){
      return [{ name: baseName(file.name) + "-pages.zip", blob: blob }];
    });
  });
};

/* ============== TOOL: DELETE PAGES ============== */

window.FileDeskTools["delete-pages"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var file = files[0];
  var toDelete = (opts.pages || "").split(",").map(function(s){ return parseInt(s.trim(),10); }).filter(function(n){ return !isNaN(n) && n>0; });
  return readAsArrayBuffer(file).then(function(buf){
    return PDFDocument.load(buf);
  }).then(function(src){
    var total = src.getPageCount();
    var keep = [];
    for (var i=0;i<total;i++){ if (toDelete.indexOf(i+1) === -1) keep.push(i); }
    if (keep.length === 0) throw new Error("That would delete every page. Leave at least one page.");
    progress(40, "Rebuilding document…");
    return PDFDocument.create().then(function(doc){
      return doc.copyPages(src, keep).then(function(pages){
        pages.forEach(function(p){ doc.addPage(p); });
        progress(85,"Saving…");
        return doc.save();
      });
    });
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-edited.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

/* ============== TOOL: ROTATE PDF ============== */

window.FileDeskTools["rotate-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var degrees = PDFLib.degrees;
  var file = files[0];
  var angle = parseInt(opts.angle || "90", 10);
  return readAsArrayBuffer(file).then(function(buf){
    return PDFDocument.load(buf);
  }).then(function(doc){
    progress(40, "Rotating pages…");
    doc.getPages().forEach(function(p){
      var current = p.getRotation().angle || 0;
      p.setRotation(degrees(current + angle));
    });
    progress(85, "Saving…");
    return doc.save();
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-rotated.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

/* ============== TOOL: COMPRESS PDF ============== */

window.FileDeskTools["compress-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var file = files[0];
  var quality = parseFloat(opts.quality || "0.7");
  return readAsArrayBuffer(file).then(function(buf){
    return pdfjsLib.getDocument({data: buf}).promise;
  }).then(function(pdf){
    return renderPagesToJpegs(pdf, 1.3, quality, progress, 10, 75);
  }).then(function(pages){
    progress(90, "Rebuilding PDF…");
    return PDFDocument.create().then(function(outDoc){
      var chain = Promise.resolve();
      pages.forEach(function(p){
        chain = chain.then(function(){
          return outDoc.embedJpg(p.dataUrl).then(function(img){
            var pg = outDoc.addPage([p.width, p.height]);
            pg.drawImage(img, {x:0, y:0, width:p.width, height:p.height});
          });
        });
      });
      return chain.then(function(){
        progress(95, "Saving…");
        return outDoc.save();
      });
    });
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-compressed.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

/* ============== TOOL: PDF TO JPG ============== */

window.FileDeskTools["pdf-to-jpg"] = function(files, opts, progress){
  var file = files[0];
  return readAsArrayBuffer(file).then(function(buf){
    return pdfjsLib.getDocument({data: buf}).promise;
  }).then(function(pdf){
    var count = pdf.numPages;
    var zip = new JSZip();
    var chain = Promise.resolve();
    for (var i=1;i<=count;i++){
      (function(pageNum){
        chain = chain.then(function(){
          progress(10 + (80*pageNum/count), "Rendering page " + pageNum + " of " + count + "…");
          return pdf.getPage(pageNum).then(function(page){
            var viewport = page.getViewport({scale: 1.8});
            var canvas = document.createElement("canvas");
            canvas.width = viewport.width; canvas.height = viewport.height;
            var ctx = canvas.getContext("2d");
            return page.render({canvasContext: ctx, viewport: viewport}).promise.then(function(){
              return new Promise(function(resolve){
                canvas.toBlob(function(blob){
                  zip.file(baseName(file.name) + "-page-" + pageNum + ".jpg", blob);
                  resolve();
                }, "image/jpeg", 0.9);
              });
            });
          });
        });
      })(i);
    }
    return chain.then(function(){
      progress(95, "Zipping…");
      return zip.generateAsync({type:"blob"});
    });
  }).then(function(blob){
    return [{ name: baseName(file.name) + "-images.zip", blob: blob }];
  });
};

/* ============== TOOL: JPG TO PDF ============== */

window.FileDeskTools["jpg-to-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  return PDFDocument.create().then(function(doc){
    var chain = Promise.resolve();
    files.forEach(function(f, i){
      chain = chain.then(function(){
        progress(10 + (75*i/files.length), "Adding " + f.name + "…");
        return readAsArrayBuffer(f).then(function(buf){
          var isPng = /png$/i.test(f.type) || /\.png$/i.test(f.name);
          return isPng ? doc.embedPng(buf) : doc.embedJpg(buf);
        }).then(function(img){
          var page = doc.addPage([img.width, img.height]);
          page.drawImage(img, {x:0,y:0,width:img.width,height:img.height});
        });
      });
    });
    return chain.then(function(){
      progress(92, "Saving…");
      return doc.save();
    });
  }).then(function(bytes){
    return [{ name:"images.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

/* ============== TOOL: PAGE NUMBERS ============== */

window.FileDeskTools["page-numbers"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var StandardFonts = PDFLib.StandardFonts;
  var rgb = PDFLib.rgb;
  var file = files[0];
  var start = parseInt(opts.start, 10);
  if (isNaN(start)) start = 1;
  var position = opts.position || "bottom-center";
  return readAsArrayBuffer(file).then(function(buf){
    return PDFDocument.load(buf);
  }).then(function(doc){
    return doc.embedFont(StandardFonts.Helvetica).then(function(font){
      var pages = doc.getPages();
      pages.forEach(function(page, i){
        progress(20 + (65*i/pages.length), "Numbering page " + (i+1) + " of " + pages.length + "…");
        var label = String(start + i);
        var size = page.getSize();
        var textWidth = font.widthOfTextAtSize(label, 10);
        var x = position === "bottom-right" ? size.width - 36 - textWidth : (size.width - textWidth) / 2;
        page.drawText(label, { x: x, y: 22, size: 10, font: font, color: rgb(0.2,0.2,0.2) });
      });
      progress(90, "Saving…");
      return doc.save();
    });
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-numbered.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

/* ============== TOOL: WATERMARK ============== */

window.FileDeskTools["watermark-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var StandardFonts = PDFLib.StandardFonts;
  var rgb = PDFLib.rgb;
  var degrees = PDFLib.degrees;
  var file = files[0];
  var text = (opts.text || "WATERMARK").trim() || "WATERMARK";
  var angleDeg = 45;
  var rad = angleDeg * Math.PI / 180;
  return readAsArrayBuffer(file).then(function(buf){
    return PDFDocument.load(buf);
  }).then(function(doc){
    return doc.embedFont(StandardFonts.HelveticaBold).then(function(font){
      var pages = doc.getPages();
      pages.forEach(function(page, i){
        progress(20 + (65*i/pages.length), "Stamping page " + (i+1) + " of " + pages.length + "…");
        var size = page.getSize();
        var fontSize = Math.max(24, Math.min(size.width, size.height) / 8);
        var textWidth = font.widthOfTextAtSize(text, fontSize);
        var cx = size.width / 2, cy = size.height / 2;
        var x = cx - (textWidth / 2) * Math.cos(rad);
        var y = cy - (textWidth / 2) * Math.sin(rad);
        page.drawText(text, {
          x: x, y: y, size: fontSize, font: font,
          color: rgb(0.6,0.6,0.6), opacity: 0.35, rotate: degrees(angleDeg)
        });
      });
      progress(90, "Saving…");
      return doc.save();
    });
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-watermarked.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

/* ============== TOOL: PROTECT PDF ============== */

window.FileDeskTools["protect-pdf"] = function(files, opts, progress){
  var file = files[0];
  var password = opts.password || "";
  if (!password) return Promise.reject(new Error("Enter a password to protect this PDF with."));
  return readAsArrayBuffer(file).then(function(buf){
    return pdfjsLib.getDocument({data: buf}).promise;
  }).then(function(pdf){
    return renderPagesToJpegs(pdf, 1.6, 0.85, progress, 10, 65);
  }).then(function(pages){
    progress(80, "Encrypting…");
    var jsPDFCtor = window.jspdf.jsPDF;
    var first = pages[0];
    var orientation = first.width > first.height ? "l" : "p";
    var doc = new jsPDFCtor({
      orientation: orientation,
      unit: "pt",
      format: [first.width, first.height],
      encryption: { userPassword: password, ownerPassword: password, userPermissions: ["print"] }
    });
    pages.forEach(function(p, i){
      if (i > 0) doc.addPage([p.width, p.height], p.width > p.height ? "l" : "p");
      doc.addImage(p.dataUrl, "JPEG", 0, 0, p.width, p.height);
    });
    var blob = doc.output("blob");
    return [{ name: baseName(file.name) + "-protected.pdf", blob: blob }];
  });
};

/* ============== TOOL: UNLOCK PDF ============== */

window.FileDeskTools["unlock-pdf"] = function(files, opts, progress){
  var file = files[0];
  var password = opts.password || "";
  return readAsArrayBuffer(file).then(function(buf){
    return pdfjsLib.getDocument({data: buf, password: password}).promise;
  }).catch(function(err){
    if (err && err.name === "PasswordException"){
      throw new Error(password ? "That password didn't open the file." : "This PDF needs a password — enter it above.");
    }
    throw err;
  }).then(function(pdf){
    return renderPagesToJpegs(pdf, 1.6, 0.85, progress, 10, 65);
  }).then(function(pages){
    progress(85, "Rebuilding without a password…");
    var PDFDocument = PDFLib.PDFDocument;
    return PDFDocument.create().then(function(doc){
      var chain = Promise.resolve();
      pages.forEach(function(p){
        chain = chain.then(function(){
          return doc.embedJpg(p.dataUrl).then(function(img){
            var pg = doc.addPage([p.width, p.height]);
            pg.drawImage(img, {x:0, y:0, width:p.width, height:p.height});
          });
        });
      });
      return chain.then(function(){ return doc.save(); });
    });
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-unlocked.pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

/* ============== TOOL: PDF TO WORD ============== */

window.FileDeskTools["pdf-to-word"] = function(files, opts, progress){
  var file = files[0];
  return readAsArrayBuffer(file).then(function(buf){
    return pdfjsLib.getDocument({data: buf}).promise;
  }).then(function(pdf){
    var count = pdf.numPages;
    var paragraphs = [];
    var chain = Promise.resolve();
    for (var i=1;i<=count;i++){
      (function(pageNum){
        chain = chain.then(function(){
          progress(10 + (75*pageNum/count), "Reading page " + pageNum + " of " + count + "…");
          return pdf.getPage(pageNum).then(function(page){
            return page.getTextContent();
          }).then(function(content){
            var lineMap = {};
            content.items.forEach(function(item){
              var y = Math.round(item.transform[5]);
              if (!lineMap[y]) lineMap[y] = [];
              lineMap[y].push(item.str);
            });
            var ys = Object.keys(lineMap).map(Number).sort(function(a,b){ return b-a; });
            ys.forEach(function(y){
              var line = lineMap[y].join(" ").replace(/\s+/g," ").trim();
              if (line) paragraphs.push(line);
            });
            paragraphs.push("");
          });
        });
      })(i);
    }
    return chain.then(function(){
      progress(90, "Building Word document…");
      return buildDocx(paragraphs);
    });
  }).then(function(blob){
    return [{ name: baseName(file.name) + ".docx", blob: blob }];
  });
};

/* ============== TOOL: WORD TO PDF ============== */

window.FileDeskTools["word-to-pdf"] = function(files, opts, progress){
  var file = files[0];
  return readAsArrayBuffer(file).then(function(buf){
    progress(25, "Reading Word document…");
    return mammoth.extractRawText({arrayBuffer: buf});
  }).then(function(result){
    progress(55, "Laying out PDF…");
    var text = result.value || "";
    var paragraphs = text.split(/\n+/);
    return buildPdfFromText(paragraphs);
  }).then(function(bytes){
    return [{ name: baseName(file.name) + ".pdf", blob:new Blob([bytes], {type:"application/pdf"}) }];
  });
};

})();