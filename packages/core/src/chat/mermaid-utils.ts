import mermaid from "mermaid";

let initialized = false;

export async function renderMermaid(code: string, id: string): Promise<string> {
  if (!initialized) {
    mermaid.initialize({
      startOnLoad: false,
      theme: "default",
      securityLevel: "loose",
      fontFamily: "var(--chat-font-mono)",
    });
    initialized = true;
  }

  try {
    const { svg } = await mermaid.render(id, code);
    return svg;
  } catch (err) {
    console.error("Mermaid rendering failed:", err);
    throw err;
  }
}

export function svgToBase64(svg: string): string {
  // Ensure the SVG has the necessary namespace
  if (!svg.includes("http://www.w3.org/2000/svg")) {
    svg = svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
  }

  // Use btoa for base64 encoding. Since SVG can contain non-ASCII characters,
  // we use a common trick to encode it as UTF-8 first.
  const utf8Encoded = encodeURIComponent(svg).replace(
    /%([0-9A-F]{2})/g,
    (_, p1) => String.fromCharCode(Number.parseInt(p1, 16)),
  );

  return btoa(utf8Encoded);
}
