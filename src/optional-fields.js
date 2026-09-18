// Remove o rótulo de RG do Word quando o campo for deixado em branco.
export function omitOptionalRg(xml, data) {
  const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const doc = new DOMParser().parseFromString(xml, "application/xml");
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
          !ranges.some(([start, end]) => offset + index >= start && offset + index < end)
        ).join("");
        offset += original.length;
      }
    }
  }
  return new XMLSerializer().serializeToString(doc);
}
