# FileDesk — Free PDF Tools

A collection of free, privacy-focused PDF tools that run entirely in your browser. No uploads, no registration, no server.

🔗 **Live Site:** [https://tools.voxcraft.site](https://tools.voxcraft.site)

---

## 🛠️ Tools Available

| Tool | Page | Description |
|------|------|-------------|
| Merge PDF | [merge-pdf.html](merge-pdf.html) | Combine multiple PDFs into one |
| Split PDF | [split-pdf.html](split-pdf.html) | Extract each page as a separate PDF |
| Compress PDF | [compress-pdf.html](compress-pdf.html) | Reduce file size (lossy) |
| PDF → Word | [pdf-to-word.html](pdf-to-word.html) | Convert PDF to editable .docx |
| Word → PDF | [word-to-pdf.html](word-to-pdf.html) | Convert .docx to PDF |
| PDF → JPG | [pdf-to-jpg.html](pdf-to-jpg.html) | Export pages as images |
| JPG → PDF | [jpg-to-pdf.html](jpg-to-pdf.html) | Turn images into a PDF |
| Rotate PDF | [rotate-pdf.html](rotate-pdf.html) | Rotate all pages |
| Delete Pages | [delete-pages.html](delete-pages.html) | Remove specific pages |
| Page Numbers | [page-numbers.html](page-numbers.html) | Add page numbers |
| Watermark | [watermark-pdf.html](watermark-pdf.html) | Add text watermark |
| Protect PDF | [protect-pdf.html](protect-pdf.html) | Password protect |
| Unlock PDF | [unlock-pdf.html](unlock-pdf.html) | Remove password |

---

## 🚀 Deployment

### Option 1: GitHub Pages (Recommended)

1. **Fork or upload** this repo to GitHub
2. Go to **Settings → Pages**
3. Select **Branch: main** and **Folder: / (root)**
4. Add your custom domain: `tools.voxcraft.site`
5. Add a `CNAME` DNS record pointing to `yourusername.github.io`

### Option 2: Netlify / Vercel

Drag and drop the folder. Zero config.

### Option 3: Any Static Host

Upload all files to your web server. No server-side processing needed.

---

## 📁 File Structure

```
├── index.html              # Homepage
├── merge-pdf.html          # Merge tool
├── split-pdf.html          # Split tool
├── compress-pdf.html       # Compress tool
├── pdf-to-word.html        # PDF to Word
├── word-to-pdf.html        # Word to PDF
├── pdf-to-jpg.html         # PDF to JPG
├── jpg-to-pdf.html         # JPG to PDF
├── rotate-pdf.html         # Rotate tool
├── delete-pages.html       # Delete pages
├── page-numbers.html       # Page numbers
├── watermark-pdf.html      # Watermark
├── protect-pdf.html        # Protect PDF
├── unlock-pdf.html         # Unlock PDF
├── tools.js                # Shared PDF processing logic
├── sitemap.xml             # Google Search Console sitemap
├── robots.txt              # Search engine instructions
└── .nojekyll               # Prevents Jekyll processing
```

---

## 📢 Ad Configuration

### AdSense (Replace when approved)
- Publisher ID: `ca-pub-3088581560119805`
- Slot IDs: Replace `YOUR_ADSENSE_SLOT_ID` in all HTML files once AdSense approves your site

### Adsterra (Already Active)
- Social Bar script is embedded in all pages
- Key: `pl30142951.effectivecpmnetwork.com`

---

## 🔍 SEO

- Unique `<title>` and `<meta description>` on every page
- Canonical URLs pointing to `tools.voxcraft.site`
- Open Graph tags for social sharing
- Schema.org `WebApplication` structured data
- Keyword-rich filenames (e.g., `merge-pdf.html`)
- `sitemap.xml` submitted to Google Search Console
- `robots.txt` for crawler instructions

---

## ⚡ Tech Stack

- **PDF-Lib** — PDF manipulation
- **PDF.js** — PDF rendering
- **JSZip** — ZIP file creation
- **Mammoth.js** — Word document parsing
- **jsPDF** — PDF generation with encryption
- **Vanilla JavaScript** — No frameworks, no build step

---

## 📄 License

MIT — Free to use, modify, and distribute.

---

Made with ❤️ in Pakistan
