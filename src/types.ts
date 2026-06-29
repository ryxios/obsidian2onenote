export type ExportMode = "single" | "folder" | "vault";

export interface OneNoteNotebook {
  id: string;
  displayName: string;
  self?: string;
}

export interface OneNoteSection {
  id: string;
  displayName: string;
  self?: string;
}

export interface OneNoteExporterSettings {
  graphAccessToken: string;
  selectedNotebookId: string;
  selectedNotebookName: string;
  selectedSectionId: string;
  selectedSectionName: string;
  defaultExportMode: ExportMode;
  cachedNotebooks: OneNoteNotebook[];
  cachedSections: OneNoteSection[];
}

export interface OneNoteExportTarget {
  notebookId: string;
  notebookName: string;
  sectionId: string;
  sectionName: string;
}

export interface ExportResult {
  filePath: string;
  success: boolean;
  message: string;
}

export const DEFAULT_SETTINGS: OneNoteExporterSettings = {
  graphAccessToken: "",
  selectedNotebookId: "",
  selectedNotebookName: "",
  selectedSectionId: "",
  selectedSectionName: "",
  defaultExportMode: "single",
  cachedNotebooks: [],
  cachedSections: []
};
