import mermaid from "mermaid";

let initialized = false;

export async function renderMermaid(code: string, id: string): Promise<string> {
  if (!initialized) {
    mermaid.initialize({
      startOnLoad: false,
      theme: "default",
      securityLevel: "loose",
      fontFamily: "var(--chat-font-mono)",
      htmlLabels: false,
      flowchart: { htmlLabels: false },
      sequence: { showSequenceNumbers: false },
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
  if (!svg.includes("http://www.w3.org/2000/svg")) {
    svg = svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
  }

  if (!svg.includes("<?xml")) {
    svg = `<?xml version="1.0" encoding="UTF-8"?>\n${svg}`;
  }

  const utf8Encoded = encodeURIComponent(svg).replace(
    /%([0-9A-F]{2})/g,
    (_, p1) => String.fromCharCode(Number.parseInt(p1, 16)),
  );

  return btoa(utf8Encoded);
}

export function svgToPngBase64(svg: string): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!svg.includes("http://www.w3.org/2000/svg")) {
      svg = svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(svg, "image/svg+xml");
    const svgElement = doc.documentElement;

    let width = parseFloat(svgElement.getAttribute("width") || "");
    let height = parseFloat(svgElement.getAttribute("height") || "");

    if (isNaN(width) || isNaN(height)) {
      const viewBox = svgElement.getAttribute("viewBox");
      if (viewBox) {
        const parts = viewBox.split(/\s+|,/);
        width = parseFloat(parts[2]);
        height = parseFloat(parts[3]);
      }
    }

    if (isNaN(width) || width <= 0) width = 800;
    if (isNaN(height) || height <= 0) height = 600;

    svgElement.style.maxWidth = "none";
    svgElement.setAttribute("width", width.toString());
    svgElement.setAttribute("height", height.toString());

    const modifiedSvg = new XMLSerializer().serializeToString(doc);

    const utf8Encoded = encodeURIComponent(modifiedSvg).replace(
      /%([0-9A-F]{2})/g,
      (_, p1) => String.fromCharCode(Number.parseInt(p1, 16)),
    );
    const b64 = btoa(utf8Encoded);
    const dataUrl = `data:image/svg+xml;base64,${b64}`;

    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      const canvas = document.createElement("canvas");
      const scale = 2;
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        reject(new Error("Failed to get canvas context"));
        return;
      }
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(scale, scale);
      ctx.drawImage(img, 0, 0, width, height);

      const pngDataUrl = canvas.toDataURL("image/png");
      const base64 = pngDataUrl.split(",")[1];
      resolve(base64);
    };
    img.onerror = () => reject(new Error("Failed to load SVG into Image"));
    img.src = dataUrl;
  });
}
