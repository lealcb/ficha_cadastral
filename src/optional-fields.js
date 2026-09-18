// Ajustes estruturais do Word antes do preenchimento pelo Docxtemplater.
// Faz duas coisas:
// 1) remove toda a seção de cônjuge quando ela não for usada;
// 2) remove o rótulo RG quando o campo opcional estiver em branco.
export function prepareTemplateXml(xml, data) {
  const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const doc = new DOMParser().parseFromString(xml, "application/xml");

  const textOf = (paragraph) =>
    Array.from(paragraph.getElementsByTagNameNS(ns, "t"))
      .map((node) => node.textContent)
      .join("");

  const paragraphs = Array.from(doc.getElementsByTagNameNS(ns, "p"));
  const start = paragraphs.find((p) => textOf(p).includes("{{#tem_conjuge}}"));
  const end = paragraphs.find((p) => textOf(p).includes("{{/tem_conjuge}}"));

  if (start && end && start.parentNode === end.parentNode) {
    if (data.tem_conjuge) {
      start.parentNode.removeChild(start);
      end.parentNode.removeChild(end);
    } else {
      let node = start;
      while (node) {
        const next = node.nextSibling;
        const isEnd = node === end;
        node.parentNode.removeChild(node);
        if (isEnd) break;
        node = next;
      }
    }
  }

  const fields = [
    ["loc_rg", /RG:\s*\{\{\s*loc_rg\s*\}\}\s*/g],
    ["conj_rg", /RG:\s*\{\{\s*conj_rg\s*\}\}\s*/g],
  ];

  for (const [key, pattern] of fields) {
    if (String(data[key] || "").trim()) continue;
    for (const paragraph of Array.from(doc.getElementsByTagNameNS(ns, "p"))) {
      const nodes = Array.from(paragraph.getElementsByTagNameNS(ns, "t"));
      const text = nodes.map((node) => node.textContent).join("");
      const ranges = Array.from(text.matchAll(pattern), (match) => [match.index, match.index + match[0].length]);
      if (!ranges.length) continue;

      let offset = 0;
      for (const node of nodes) {
        const original = node.textContent;
        node.textContent = original.split("").filter((_, index) =>
          !ranges.some(([from, to]) => offset + index >= from && offset + index < to)
        ).join("");
        offset += original.length;
      }
    }
  }

  return new XMLSerializer().serializeToString(doc);
}
