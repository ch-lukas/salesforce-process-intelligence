// Writes a colour and font theme into a .pptx built by pptxgenjs.
// pptxgenjs can reference scheme colours (accent1, tx2, ...) but can't set their values,
// so without this step the deck would render in Office's stock palette.
const fs = require("fs");
const JSZip = require("jszip");

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const escapeXml = s => String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" }[c]));

async function applyTheme(file, theme) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const parts = Object.keys(zip.files).filter(n => /^ppt\/theme\/theme\d+\.xml$/.test(n));
  for (const part of parts) {
    let xml = await zip.file(part).async("string");
    for (const slot of SLOTS) {
      const hex = theme.colors[slot];
      if (!/^[0-9A-Fa-f]{6}$/.test(hex || "")) throw new Error(`theme.colors.${slot} must be a 6-digit hex colour`);
      xml = xml.replace(new RegExp(`<a:${slot}>[\\s\\S]*?</a:${slot}>`), `<a:${slot}><a:srgbClr val="${hex.toUpperCase()}"/></a:${slot}>`);
    }
    const name = escapeXml(theme.name);
    xml = xml
      .replace(/(<a:theme\b[^>]*\bname=")[^"]*"/, `$1${name}"`)
      .replace(/(<a:clrScheme\b[^>]*\bname=")[^"]*"/, `$1${name}"`)
      .replace(/(<a:fontScheme\b[^>]*\bname=")[^"]*"/, `$1${name}"`)
      .replace(/(<a:majorFont>\s*<a:latin typeface=")[^"]*"/, `$1${escapeXml(theme.headFontFace)}"`)
      .replace(/(<a:minorFont>\s*<a:latin typeface=")[^"]*"/, `$1${escapeXml(theme.bodyFontFace)}"`);
    zip.file(part, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

module.exports = { applyTheme };
