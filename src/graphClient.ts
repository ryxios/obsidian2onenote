import { Notice } from "obsidian";
import { OneNoteNotebook, OneNoteSection } from "./types";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

interface GraphCollection<T> {
  value: T[];
}

export class GraphClient {
  constructor(private readonly getAccessToken: () => string) {}

  async getNotebooks(): Promise<OneNoteNotebook[]> {
    const data = await this.request<GraphCollection<OneNoteNotebook>>("/me/onenote/notebooks");
    return data.value.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }

  async getSections(notebookId: string): Promise<OneNoteSection[]> {
    if (!notebookId) {
      return [];
    }
    const data = await this.request<GraphCollection<OneNoteSection>>(`/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sections`);
    return data.value.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }

  async createPage(html: string, sectionId: string): Promise<void> {
    if (!sectionId) {
      throw new Error("Select a OneNote section before exporting.");
    }

    await this.request<void>(`/me/onenote/pages?sectionId=${encodeURIComponent(sectionId)}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/xhtml+xml"
      },
      body: html
    });
  }

  private async request<T>(path: string, init: RequestInit = {}): Promise<T> {
    const token = this.getAccessToken().trim();
    if (!token) {
      throw new Error("Microsoft Graph access token is missing. Add it in OneNote Exporter settings.");
    }

    const response = await fetch(`${GRAPH_ROOT}${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ...(init.headers ?? {})
      }
    });

    if (response.status === 401) {
      new Notice("Microsoft Graph token expired or unauthorized. Update the token in settings.");
      throw new Error("Microsoft Graph returned 401 Unauthorized. Refresh your access token.");
    }

    if (!response.ok) {
      const details = await response.text();
      throw new Error(`Microsoft Graph request failed (${response.status} ${response.statusText}): ${details}`);
    }

    if (response.status === 204) {
      return undefined as T;
    }

    const text = await response.text();
    return text ? (JSON.parse(text) as T) : (undefined as T);
  }
}
