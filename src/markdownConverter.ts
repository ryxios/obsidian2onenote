import MarkdownIt from "markdown-it";
import { App, TFile } from "obsidian";

export class MarkdownConverter {
  private readonly markdown: MarkdownIt;

  constructor(private readonly app: App) {
    this.markdown = new MarkdownIt({ html: false, linkify: true, breaks: true });
  }

  async convert(markdown: string, sourceFile: TFile): Promise<string> {
    const sanitized = await this.prepareMarkdown(markdown, sourceFile);
    const body = this.markdown.render(sanitized);
    const title = this.escapeXml(sourceFile.basename);

    return `<!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml"><head><title>${title}</title><meta name="created" content="${new Date().toISOString()}" /></head><body>${body}</body></html>`;
  }

  private async prepareMarkdown(markdown: string, sourceFile: TFile): Promise<string> {
    let content = markdown;
    content = content.replace(/```dataview[\s\S]*?```/gi, "");
    content = content.replace(/```dataviewjs[\s\S]*?```/gi, "");
    content = content.replace(/^>\s*\[!([\w-]+)\][^\n]*(\n>.*)*/gim, (match) => match.replace(/^>\s*\[![^\]]+\]\s*/im, "> "));
    content = await this.replaceObsidianEmbeds(content, sourceFile);
    content = content.replace(/!\[\[[^\]]+\]\]/g, "");
    content = content.replace(/\[\[([^\]|]+)\|([^\]]+)\]\]/g, "$2");
    content = content.replace(/\[\[([^\]]+)\]\]/g, "$1");
    return content;
  }

  private async replaceObsidianEmbeds(content: string, sourceFile: TFile): Promise<string> {
    const matches = Array.from(content.matchAll(/!\[\[([^\]|#]+)(?:#[^\]|]+)?(?:\|[^\]]+)?\]\]/g));
    let output = content;

    for (const match of matches) {
      const embedName = match[1].trim();
      const linked = this.app.metadataCache.getFirstLinkpathDest(embedName, sourceFile.path);
      if (!linked || !this.isImage(linked.extension)) {
        output = output.replace(match[0], "");
        continue;
      }

      try {
        const dataUrl = await this.fileToDataUrl(linked);
        output = output.replace(match[0], `![${linked.basename}](${dataUrl})`);
      } catch {
        output = output.replace(match[0], "");
      }
    }

    return output;
  }

  private async fileToDataUrl(file: TFile): Promise<string> {
    const bytes = await this.app.vault.readBinary(file);
    const binary = Array.from(new Uint8Array(bytes), (byte) => String.fromCharCode(byte)).join("");
    return `data:${this.mimeType(file.extension)};base64,${btoa(binary)}`;
  }

  private isImage(extension: string): boolean {
    return ["png", "jpg", "jpeg", "gif", "svg", "webp", "bmp"].includes(extension.toLowerCase());
  }

  private mimeType(extension: string): string {
    const ext = extension.toLowerCase();
    if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
    if (ext === "svg") return "image/svg+xml";
    if (ext === "webp") return "image/webp";
    if (ext === "gif") return "image/gif";
    if (ext === "bmp") return "image/bmp";
    return "image/png";
  }

  private escapeXml(value: string): string {
    return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
}
