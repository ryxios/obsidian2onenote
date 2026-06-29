import { App, Notice, TFile, TFolder } from "obsidian";
import { GraphClient } from "./graphClient";
import { MarkdownConverter } from "./markdownConverter";
import { ExportResult } from "./types";

export class OneNoteExporter {
  constructor(
    private readonly app: App,
    private readonly graphClient: GraphClient,
    private readonly converter: MarkdownConverter,
    private readonly getSectionId: () => string
  ) {}

  async exportFile(file: TFile): Promise<ExportResult> {
    new Notice(`Exporting ${file.basename} to OneNote...`);
    try {
      const markdown = await this.app.vault.read(file);
      const html = await this.converter.convert(markdown, file);
      await this.graphClient.createPage(html, this.getSectionId());
      console.info(`[OneNote Exporter] Exported ${file.path}`);
      new Notice(`Exported ${file.basename} to OneNote.`);
      return { filePath: file.path, success: true, message: "Exported successfully" };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error(`[OneNote Exporter] Failed to export ${file.path}: ${message}`, error);
      new Notice(`Failed to export ${file.basename}: ${message}`);
      return { filePath: file.path, success: false, message };
    }
  }

  async exportFolder(folder: TFolder): Promise<ExportResult[]> {
    const files = this.getMarkdownFiles(folder).sort((a, b) => a.path.localeCompare(b.path));
    return this.exportBatch(files, `folder ${folder.path || "/"}`);
  }

  async exportVault(): Promise<ExportResult[]> {
    const files = this.app.vault.getMarkdownFiles().sort((a, b) => a.path.localeCompare(b.path));
    return this.exportBatch(files, "entire vault");
  }

  private async exportBatch(files: TFile[], label: string): Promise<ExportResult[]> {
    const total = files.length;
    const results: ExportResult[] = [];
    new Notice(`OneNote export started for ${label}: ${total} note(s).`);

    for (let index = 0; index < files.length; index++) {
      const file = files[index];
      new Notice(`OneNote export ${index + 1}/${total}: ${file.path}`);
      results.push(await this.exportFile(file));
    }

    const successes = results.filter((result) => result.success).length;
    const failures = total - successes;
    new Notice(`OneNote export complete: ${successes} succeeded, ${failures} failed.`);
    console.table(results);
    return results;
  }

  private getMarkdownFiles(folder: TFolder): TFile[] {
    const files: TFile[] = [];
    for (const child of folder.children) {
      if (child instanceof TFile && child.extension === "md") {
        files.push(child);
      } else if (child instanceof TFolder) {
        files.push(...this.getMarkdownFiles(child));
      }
    }
    return files;
  }
}
