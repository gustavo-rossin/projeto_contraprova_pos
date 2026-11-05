// Grava as cores do tema no PowerPoint gerado pelo pptxgenjs.
// O pptxgenjs define as fontes do tema, mas mantém a paleta padrão do Office; este passo
// substitui as 12 cores do esquema (dk1, lt1, ..., accent6, hlink, folHlink) em ppt/theme/theme1.xml.
const fs = require("fs");
const JSZip = require("jszip");

async function aplicarTema(arquivo, tema) {
  const zip = await JSZip.loadAsync(fs.readFileSync(arquivo));
  const temas = Object.keys(zip.files).filter(n => /^ppt\/theme\/theme\d+\.xml$/.test(n));
  for (const caminho of temas) {
    let xml = await zip.file(caminho).async("string");
    for (const [nome, hex] of Object.entries(tema.colors)) {
      if (!/^[0-9A-Fa-f]{6}$/.test(hex)) throw new Error(`Cor inválida para ${nome}: ${hex}`);
      const re = new RegExp(`(<a:${nome}>)[\\s\\S]*?(</a:${nome}>)`);
      xml = xml.replace(re, `$1<a:srgbClr val="${hex.toUpperCase()}"/>$2`);
    }
    const nome = String(tema.name).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    xml = xml.replace(/<a:clrScheme name="[^"]*"/, `<a:clrScheme name="${nome}"`);
    zip.file(caminho, xml);
  }
  fs.writeFileSync(arquivo, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

module.exports = { aplicarTema };
