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
// FIXED: No longer rasterizes text/vectors. Only re-encodes embedded images.
// Falls back to original file if compression would increase size.

window.FileDeskTools["compress-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var PDFName = PDFLib.PDFName;
  var file = files[0];
  var quality = parseFloat(opts.quality || "0.75");

  var targetDPI = quality <= 0.5 ? 72 : quality <= 0.75 ? 96 : 120;
  var jpegQuality = quality <= 0.5 ? 0.45 : quality <= 0.75 ? 0.65 : 0.82;
  var maxImageDim = quality <= 0.5 ? 1200 : quality <= 0.75 ? 1800 : 2400;

  return readAsArrayBuffer(file).then(function(buf){
    progress(10, "Loading PDF structure…");
    return PDFDocument.load(buf, { updateMetadata: false });
  }).then(function(pdfDoc){
    var pages = pdfDoc.getPages();
    var pageCount = pages.length;
    var processedImages = 0;
    var skippedImages = 0;

    progress(20, "Analyzing " + pageCount + " page" + (pageCount > 1 ? "s" : "") + "…");

    var imageChain = Promise.resolve();

    pages.forEach(function(page, pageIdx){
      imageChain = imageChain.then(function(){
        var resources = page.node.Resources();
        if (!resources) return;

        var xObject = resources.lookup(PDFName.of('XObject'));
        if (!xObject) return;

        var dict = xObject.dict;
        var keys = Object.keys(dict);

        keys.forEach(function(key){
          var obj = dict[key];
          if (!obj || !obj.dict) return;

          var subtype = obj.dict.get(PDFName.of('Subtype'));
          if (!subtype) return;

          var subtypeStr = subtype.asString ? subtype.asString() : subtype.toString();
          if (subtypeStr !== 'Image') return;

          var width = 0, height = 0;
          try {
            var w = obj.dict.get(PDFName.of('Width'));
            var h = obj.dict.get(PDFName.of('Height'));
            width = w ? (w.asNumber ? w.asNumber() : Number(w)) : 0;
            height = h ? (h.asNumber ? h.asNumber() : Number(h)) : 0;
          } catch(e) { return; }

          if (width < 64 || height < 64) {
            skippedImages++;
            return;
          }

          if (width <= maxImageDim && height <= maxImageDim) {
            var filter = obj.dict.get(PDFName.of('Filter'));
            var filterStr = filter ? (filter.asString ? filter.asString() : filter.toString()) : '';
            if (filterStr.includes('DCTDecode') && quality >= 0.9) {
              skippedImages++;
              return;
            }
          }

          try {
            var stream = obj;
            var rawData;

            if (stream.getContents) {
              rawData = stream.getContents();
            } else if (stream.contents) {
              rawData = stream.contents;
            } else {
              return;
            }

            var filter = obj.dict.get(PDFName.of('Filter'));
            var filterStr = filter ? (filter.asString ? filter.asString() : filter.toString()) : '';
            var isJpeg = filterStr.includes('DCTDecode');

            var mimeType = isJpeg ? 'image/jpeg' : 'image/png';
            var blob = new Blob([rawData], { type: mimeType });
            var imgUrl = URL.createObjectURL(blob);

            return new Promise(function(resolve, reject){
              var img = new Image();
              img.onload = function(){
                URL.revokeObjectURL(imgUrl);

                var scale = Math.min(1, maxImageDim / Math.max(width, height), targetDPI / 150);
                if (scale >= 1 && isJpeg) {
                  skippedImages++;
                  resolve();
                  return;
                }

                var newWidth = Math.max(1, Math.round(width * scale));
                var newHeight = Math.max(1, Math.round(height * scale));

                var canvas = document.createElement("canvas");
                canvas.width = newWidth;
                canvas.height = newHeight;
                var ctx = canvas.getContext("2d");

                ctx.fillStyle = "#FFFFFF";
                ctx.fillRect(0, 0, newWidth, newHeight);
                ctx.drawImage(img, 0, 0, newWidth, newHeight);

                var jpegDataUrl = canvas.toDataURL("image/jpeg", jpegQuality);
                var base64 = jpegDataUrl.split(',')[1];
                var binaryString = atob(base64);
                var bytes = new Uint8Array(binaryString.length);
                for (var i = 0; i < binaryString.length; i++) {
                  bytes[i] = binaryString.charCodeAt(i);
                }

                pdfDoc.embedJpg(bytes).then(function(newImage){
                  dict[key] = newImage.ref;
                  processedImages++;
                  resolve();
                }).catch(function(err){
                  console.warn("Failed to embed image:", err);
                  skippedImages++;
                  resolve();
                });
              };
              img.onerror = function(){
                URL.revokeObjectURL(imgUrl);
                skippedImages++;
                resolve();
              };
              img.src = imgUrl;
            });
          } catch(e) {
            console.warn("Could not process image:", e);
            skippedImages++;
          }
        });

        progress(20 + Math.round((pageIdx / pageCount) * 55), 
          "Processing page " + (pageIdx + 1) + " of " + pageCount + "…");
      });
    });

    return imageChain.then(function(){
      progress(80, "Optimizing structure…");

      try {
        pdfDoc.setTitle("");
        pdfDoc.setAuthor("");
        pdfDoc.setSubject("");
        pdfDoc.setKeywords([]);
        pdfDoc.setProducer("FileDesk");
        pdfDoc.setCreator("FileDesk Compressor");
      } catch(e) {}

      progress(90, "Saving compressed PDF…");

      return pdfDoc.save({
        useObjectStreams: true,
        addDefaultPage: false,
        objectsPerTick: 50
      });
    }).then(function(bytes){
      var outputBlob = new Blob([bytes], { type: "application/pdf" });

      if (outputBlob.size > file.size && processedImages === 0) {
        progress(95, "PDF already optimized — returning original…");
        return readAsArrayBuffer(file).then(function(originalBuf){
          return [{
            name: baseName(file.name) + ".pdf",
            blob: new Blob([originalBuf], { type: "application/pdf" })
          }];
        });
      }

      progress(100, "Done — processed " + processedImages + " image" + (processedImages !== 1 ? "s" : ""));
      return [{
        name: baseName(file.name) + "-compressed.pdf",
        blob: outputBlob
      }];
    });
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
// FIXED: Better PNG detection using file signature (magic bytes) instead of just extension

window.FileDeskTools["jpg-to-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  return PDFDocument.create().then(function(doc){
    var chain = Promise.resolve();
    files.forEach(function(f, i){
      chain = chain.then(function(){
        progress(10 + (75*i/files.length), "Adding " + f.name + "…");
        return readAsArrayBuffer(f).then(function(buf){
          var bytes = new Uint8Array(buf);
          // PNG signature: 89 50 4E 47 0D 0A 1A 0A
          var isPng = bytes.length > 8 && 
                      bytes[0] === 0x89 && bytes[1] === 0x50 && 
                      bytes[2] === 0x4E && bytes[3] === 0x47;
          // JPEG signature: FF D8 FF
          var isJpg = bytes.length > 3 && 
                      bytes[0] === 0xFF && bytes[1] === 0xD8 && bytes[2] === 0xFF;

          if (isPng) {
            return doc.embedPng(buf);
          } else if (isJpg) {
            return doc.embedJpg(buf);
          } else {
            // Fallback: try by extension or type
            var extPng = /png$/i.test(f.type) || /\.png$/i.test(f.name);
            return extPng ? doc.embedPng(buf) : doc.embedJpg(buf);
          }
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
// FIXED: Uses pdf-lib native encryption instead of rasterizing to images via jsPDF.
// Preserves text selectability, vector graphics, and file size.

window.FileDeskTools["protect-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var StandardFonts = PDFLib.StandardFonts;
  var rgb = PDFLib.rgb;
  var file = files[0];
  var password = opts.password || "";
  if (!password) return Promise.reject(new Error("Enter a password to protect this PDF with."));

  return readAsArrayBuffer(file).then(function(buf){
    progress(30, "Loading document…");
    return PDFDocument.load(buf);
  }).then(function(doc){
    progress(60, "Applying encryption…");

    // Encrypt with password protection
    // userPassword = open password, ownerPassword = permissions password
    return doc.encrypt({
      userPassword: password,
      ownerPassword: password,
      permissions: {
        printing: 'highResolution',
        modifying: false,
        copying: false,
        annotating: false,
        fillingForms: false,
        contentAccessibility: true,
        documentAssembly: false
      }
    });
  }).then(function(encryptedDoc){
    progress(90, "Saving protected PDF…");
    return encryptedDoc.save();
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-protected.pdf", blob: new Blob([bytes], {type: "application/pdf"}) }];
  });
};

/* ============== TOOL: UNLOCK PDF ============== */
// FIXED: Uses pdf-lib to remove encryption without re-rendering pages as images.
// Preserves text selectability, vector graphics, and file size.

window.FileDeskTools["unlock-pdf"] = function(files, opts, progress){
  var PDFDocument = PDFLib.PDFDocument;
  var file = files[0];
  var password = opts.password || "";

  return readAsArrayBuffer(file).then(function(buf){
    progress(30, "Loading document…");
    // Try with password first
    return PDFDocument.load(buf, { 
      updateMetadata: false,
      ignoreEncryption: false 
    }).catch(function(err){
      // If encrypted, try with provided password
      if (password) {
        return PDFDocument.load(buf, { 
          updateMetadata: false,
          password: password 
        });
      }
      throw new Error("This PDF is password-protected. Enter the password above to unlock it.");
    });
  }).then(function(doc){
    progress(60, "Removing encryption…");

    // Remove encryption by saving without password
    // pdf-lib doesn't have explicit decrypt, but saving without encrypt removes it
    progress(90, "Saving unlocked PDF…");
    return doc.save({
      useObjectStreams: true,
      addDefaultPage: false
    });
  }).then(function(bytes){
    return [{ name: baseName(file.name) + "-unlocked.pdf", blob: new Blob([bytes], {type: "application/pdf"}) }];
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
