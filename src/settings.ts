import { App, Notice, PluginSettingTab, Setting } from "obsidian";
import OneNoteExporterPlugin from "./main";
import { ExportMode } from "./types";

export class OneNoteExporterSettingTab extends PluginSettingTab {
  constructor(app: App, private readonly plugin: OneNoteExporterPlugin) {
    super(app, plugin);
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "Obsidian OneNote Exporter" });

    new Setting(containerEl)
      .setName("Graph Access Token")
      .setDesc("Microsoft Graph bearer token with OneNote permissions.")
      .addText((text) => text
        .setPlaceholder("eyJ0eXAiOiJKV1Qi...")
        .setValue(this.plugin.settings.graphAccessToken)
        .onChange(async (value) => {
          this.plugin.settings.graphAccessToken = value;
          await this.plugin.saveSettings();
        }));

    new Setting(containerEl)
      .setName("Notebook")
      .setDesc("Choose the target OneNote notebook. Use Refresh Notebooks if this list is empty.")
      .addDropdown((dropdown) => {
        dropdown.addOption("", "Select a notebook");
        for (const notebook of this.plugin.settings.cachedNotebooks) {
          dropdown.addOption(notebook.id, notebook.displayName);
        }
        dropdown.setValue(this.plugin.settings.selectedNotebookId);
        dropdown.onChange(async (value) => {
          const notebook = this.plugin.settings.cachedNotebooks.find((item) => item.id === value);
          this.plugin.settings.selectedNotebookId = value;
          this.plugin.settings.selectedNotebookName = notebook?.displayName ?? "";
          this.plugin.settings.selectedSectionId = "";
          this.plugin.settings.selectedSectionName = "";
          await this.plugin.saveSettings();
          if (value) await this.plugin.refreshSections();
          this.display();
        });
      });

    new Setting(containerEl)
      .setName("Section")
      .setDesc("Choose the target OneNote section for newly created pages.")
      .addDropdown((dropdown) => {
        dropdown.addOption("", "Select a section");
        for (const section of this.plugin.settings.cachedSections) {
          dropdown.addOption(section.id, section.displayName);
        }
        dropdown.setValue(this.plugin.settings.selectedSectionId);
        dropdown.onChange(async (value) => {
          const section = this.plugin.settings.cachedSections.find((item) => item.id === value);
          this.plugin.settings.selectedSectionId = value;
          this.plugin.settings.selectedSectionName = section?.displayName ?? "";
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName("Default export mode")
      .addDropdown((dropdown) => dropdown
        .addOption("single", "Single note")
        .addOption("folder", "Folder batch")
        .addOption("vault", "Full vault batch")
        .setValue(this.plugin.settings.defaultExportMode)
        .onChange(async (value) => {
          this.plugin.settings.defaultExportMode = value as ExportMode;
          await this.plugin.saveSettings();
        }));

    new Setting(containerEl)
      .setName("Refresh OneNote data")
      .setDesc("Load notebooks and sections from Microsoft Graph.")
      .addButton((button) => button.setButtonText("Refresh Notebooks").onClick(async () => {
        await this.plugin.refreshNotebooks();
        this.display();
        new Notice("Notebooks refreshed.");
      }))
      .addButton((button) => button.setButtonText("Refresh Sections").onClick(async () => {
        await this.plugin.refreshSections();
        this.display();
        new Notice("Sections refreshed.");
      }));
  }
}
