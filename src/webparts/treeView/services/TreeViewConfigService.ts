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

  /**
   * Limpa o cache persistente (PublishedTreeData) do registro da página.
   * Isso força o componente a buscar novos dados do SharePoint na próxima renderização.
   */
  public static async clearPublishedData(pageUrl: string): Promise<void> {
    const existing = await this.loadByPage(pageUrl);

    if (existing?.Id) {
      await pnp.sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(existing.Id)
        .update({
          PublishedTreeData: null,
        });
      console.log(
        `[TreeViewConfigService] Cache persistente (PublishedTreeData) limpo para a página: ${pageUrl}`
      );
    } else {
      // Se não houver registro, não há o que limpar, mas podemos garantir que as configs mínimas existam.
      console.log(
        `[TreeViewConfigService] Nenhum registro encontrado para limpar o cache persistente: ${pageUrl}`
      );
    }
  }
}
