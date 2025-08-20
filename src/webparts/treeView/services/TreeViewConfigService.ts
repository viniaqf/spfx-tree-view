import pnp from "sp-pnp-js";

export interface ITreeViewConfigRecord {
  Id?: number;
  PageURL: string;
  Library?: string;
  Hierarchy?: string;
  PublishedTreeData?: string;
}

export default class TreeViewConfigService {
  private static readonly listTitle = "TreeViewConfigs";

  // Retorna o caminho (server-relative) da página atual.
  public static getCurrentPageUrl(): string {
    return window.location.pathname;
  }

  // Busca 1 registro da página (ou null).
  public static async loadByPage(
    pageUrl: string
  ): Promise<ITreeViewConfigRecord | null> {
    const items = await pnp.sp.web.lists
      .getByTitle(this.listTitle)
      .items.filter(`PageURL eq '${pageUrl.replace(/'/g, "''")}'`)
      .top(1)
      .get();

    if (!items || items.length === 0) return null;
    const it = items[0];
    return {
      Id: it.Id,
      PageURL: it.PageURL,
      Library: it.Library,
      Hierarchy: it.Hierarchy,
      PublishedTreeData: it.PublishedTreeData,
    };
  }

  // Upsert por PageURL.
  public static async upsert(rec: ITreeViewConfigRecord): Promise<void> {
    const existing = await this.loadByPage(rec.PageURL);

    const payload: any = {
      PageURL: rec.PageURL,
      Library: rec.Library ?? null,
      Hierarchy: rec.Hierarchy ?? null,
      PublishedTreeData: rec.PublishedTreeData ?? null,
    };

    if (existing?.Id) {
      await pnp.sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(existing.Id)
        .update(payload);
    } else {
      await pnp.sp.web.lists.getByTitle(this.listTitle).items.add(payload);
    }
  }

  // Apenas atualiza o campo PublishedTreeData do registro da página. Cria se não existir.
  public static async upsertPublishedData(
    pageUrl: string,
    jsonAllItems: string,
    library?: string,
    hierarchy?: string
  ): Promise<void> {
    await this.upsert({
      PageURL: pageUrl,
      Library: library,
      Hierarchy: hierarchy,
      PublishedTreeData: jsonAllItems,
    });
  }
}
