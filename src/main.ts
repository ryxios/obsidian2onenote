import { FuzzyMatch, FuzzySuggestModal, Notice, Plugin, TFile, TFolder } from "obsidian";
import { OneNoteExporter } from "./exporter";
import { GraphClient } from "./graphClient";
import { MarkdownConverter } from "./markdownConverter";
import { OneNoteExporterSettingTab } from "./settings";
import { DEFAULT_SETTINGS, OneNoteExportTarget, OneNoteExporterSettings, OneNoteNotebook, OneNoteSection } from "./types";

export default class OneNoteExporterPlugin extends Plugin {
  settings: OneNoteExporterSettings = DEFAULT_SETTINGS;
  private graphClient!: GraphClient;
  private converter!: MarkdownConverter;
  private exporter!: OneNoteExporter;

  async onload(): Promise<void> {
    await this.loadSettings();
    this.graphClient = new GraphClient(() => this.settings.graphAccessToken);
    this.converter = new MarkdownConverter(this.app);
    this.exporter = new OneNoteExporter(this.app, this.graphClient, this.converter);

    this.addSettingTab(new OneNoteExporterSettingTab(this.app, this));
    this.addRibbonIcon("upload", "Export current note to OneNote", () => this.exportCurrentNote());

    this.addCommand({ id: "export-current-note-to-onenote", name: "Export current note to OneNote", callback: () => this.exportCurrentNote() });
    this.addCommand({ id: "export-folder-to-onenote", name: "Export folder to OneNote", callback: () => this.exportFolder() });
    this.addCommand({ id: "export-vault-to-onenote", name: "Export vault to OneNote", callback: () => this.exportVault() });
    this.addCommand({ id: "refresh-onenote-notebooks", name: "Refresh Notebooks", callback: () => this.refreshNotebooks() });
    this.addCommand({ id: "refresh-onenote-sections", name: "Refresh Sections", callback: () => this.refreshSections() });
  }

  async loadSettings(): Promise<void> {
    this.settings = { ...DEFAULT_SETTINGS, ...(await this.loadData()) };
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  async refreshNotebooks(): Promise<void> {
    try {
      this.settings.cachedNotebooks = await this.graphClient.getNotebooks();
      if (this.settings.selectedNotebookId && !this.settings.cachedNotebooks.some((notebook) => notebook.id === this.settings.selectedNotebookId)) {
        this.settings.selectedNotebookId = "";
        this.settings.selectedNotebookName = "";
        this.settings.selectedSectionId = "";
        this.settings.selectedSectionName = "";
        this.settings.cachedSections = [];
      }
      await this.saveSettings();
      new Notice(`Loaded ${this.settings.cachedNotebooks.length} OneNote notebook(s).`);
    } catch (error) {
      this.handleError("Failed to refresh OneNote notebooks", error);
    }
  }

  async refreshSections(): Promise<void> {
    if (!this.settings.selectedNotebookId) {
      new Notice("Select a OneNote notebook before refreshing sections.");
      return;
    }

    try {
      this.settings.cachedSections = await this.graphClient.getSections(this.settings.selectedNotebookId);
      if (this.settings.selectedSectionId && !this.settings.cachedSections.some((section) => section.id === this.settings.selectedSectionId)) {
        this.settings.selectedSectionId = "";
        this.settings.selectedSectionName = "";
      }
      await this.saveSettings();
      new Notice(`Loaded ${this.settings.cachedSections.length} OneNote section(s).`);
    } catch (error) {
      this.handleError("Failed to refresh OneNote sections", error);
    }
  }

  private async exportCurrentNote(): Promise<void> {
    if (!this.ensureReady()) return;
    const file = this.app.workspace.getActiveFile();
    if (!file) {
      new Notice("Open a Markdown note before exporting to OneNote.");
      return;
    }
    const target = await this.chooseExportTarget();
    if (!target) return;
    await this.exporter.exportFile(file, target);
  }

  private async exportFolder(): Promise<void> {
    if (!this.ensureReady()) return;
    const folders = this.getFolders();
    if (!folders.length) {
      new Notice("No folders found in this vault.");
      return;
    }
    const target = await this.chooseExportTarget();
    if (!target) return;
    new FolderSuggestModal(this.app, folders, async (folder) => this.exporter.exportFolder(folder, target)).open();
  }

  private async exportVault(): Promise<void> {
    if (!this.ensureReady()) return;
    const target = await this.chooseExportTarget();
    if (!target) return;
    await this.exporter.exportVault(target);
  }

  private ensureReady(): boolean {
    if (!this.settings.graphAccessToken.trim()) {
      new Notice("Add a Microsoft Graph access token in OneNote Exporter settings.");
      return false;
    }
    return true;
  }

  private async chooseExportTarget(): Promise<OneNoteExportTarget | null> {
    try {
      const notebooks = await this.getAvailableNotebooks();
      if (!notebooks.length) {
        new Notice("No OneNote notebooks found. Refresh notebooks or check your Microsoft Graph permissions.");
        return null;
      }

      const notebook = await this.openSuggestModal(new NotebookSuggestModal(this.app, notebooks));
      if (!notebook) return null;

      const sections = await this.graphClient.getSections(notebook.id);
      this.settings.selectedNotebookId = notebook.id;
      this.settings.selectedNotebookName = notebook.displayName;
      this.settings.cachedSections = sections;
      this.settings.selectedSectionId = "";
      this.settings.selectedSectionName = "";

      if (!sections.length) {
        await this.saveSettings();
        new Notice(`No sections found in notebook ${notebook.displayName}.`);
        return null;
      }

      const section = await this.openSuggestModal(new SectionSuggestModal(this.app, notebook, sections));
      if (!section) {
        await this.saveSettings();
        return null;
      }

      this.settings.selectedSectionId = section.id;
      this.settings.selectedSectionName = section.displayName;
      await this.saveSettings();

      return {
        notebookId: notebook.id,
        notebookName: notebook.displayName,
        sectionId: section.id,
        sectionName: section.displayName
      };
    } catch (error) {
      this.handleError("Failed to choose OneNote export target", error);
      return null;
    }
  }

  private async getAvailableNotebooks(): Promise<OneNoteNotebook[]> {
    if (!this.settings.cachedNotebooks.length) {
      this.settings.cachedNotebooks = await this.graphClient.getNotebooks();
      await this.saveSettings();
    }
    return this.settings.cachedNotebooks;
  }

  private openSuggestModal<T>(modal: FuzzySuggestModal<T>): Promise<T | null> {
    return new Promise((resolve) => {
      const originalOnClose = modal.onClose.bind(modal);
      let resolved = false;
      modal.onClose = () => {
        originalOnClose();
        if (!resolved) {
          resolved = true;
          resolve(null);
        }
      };
      const originalOnChooseItem = modal.onChooseItem.bind(modal);
      modal.onChooseItem = (item, evt) => {
        originalOnChooseItem(item, evt);
        if (!resolved) {
          resolved = true;
          resolve(item);
        }
      };
      modal.open();
    });
  }

  private getFolders(): TFolder[] {
    const folders: TFolder[] = [];
    const visit = (folder: TFolder) => {
      folders.push(folder);
      for (const child of folder.children) {
        if (child instanceof TFolder) visit(child);
      }
    };
    visit(this.app.vault.getRoot());
    return folders;
  }

  private handleError(prefix: string, error: unknown): void {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[OneNote Exporter] ${prefix}: ${message}`, error);
    new Notice(`${prefix}: ${message}`);
  }
}

class NotebookSuggestModal extends FuzzySuggestModal<OneNoteNotebook> {
  constructor(app: OneNoteExporterPlugin["app"], private readonly notebooks: OneNoteNotebook[]) {
    super(app);
    this.setPlaceholder("Select the OneNote notebook for this export");
  }

  getItems(): OneNoteNotebook[] {
    return this.notebooks;
  }

  getItemText(notebook: OneNoteNotebook): string {
    return notebook.displayName;
  }

  onChooseItem(): void {
    return;
  }
}

class SectionSuggestModal extends FuzzySuggestModal<OneNoteSection> {
  constructor(app: OneNoteExporterPlugin["app"], private readonly notebook: OneNoteNotebook, private readonly sections: OneNoteSection[]) {
    super(app);
    this.setPlaceholder(`Select a section in ${notebook.displayName}`);
  }

  getItems(): OneNoteSection[] {
    return this.sections;
  }

  getItemText(section: OneNoteSection): string {
    return `${this.notebook.displayName} / ${section.displayName}`;
  }

  onChooseItem(): void {
    return;
  }
}

class FolderSuggestModal extends FuzzySuggestModal<TFolder> {
  constructor(app: OneNoteExporterPlugin["app"], private readonly folders: TFolder[], private readonly onChoose: (folder: TFolder) => void) {
    super(app);
    this.setPlaceholder("Select a folder to export to OneNote");
  }

  getItems(): TFolder[] {
    return this.folders;
  }

  getItemText(folder: TFolder): string {
    return folder.path || "/";
  }

  onChooseItem(folder: TFolder): void {
    this.onChoose(folder);
  }
}
