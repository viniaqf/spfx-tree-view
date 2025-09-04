import * as React from 'react';
import styles from '../components/TreeView.module.scss';
import { ITreeViewProps } from './ITreeViewProps';
import pnp from "sp-pnp-js";
import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from 'office-ui-fabric-react';
import { getTranslations, getUserLanguage } from '../../../utils/getTranslations';
import IframePreview from './IframePreview';
import SplitterLayout from 'react-splitter-layout';
import 'react-splitter-layout/lib/index.css';
import TreeViewConfigService from '../services/TreeViewConfigService';
import { injectCssStringOnce } from '../../../utils/localCssInjector';
import { HIDE_SWITCHER_CSS } from '../../../styles/spfx_style';




interface ITreeNode {
  key: string;
  label: string;
  icon?: string;
  url?: string;
  isFolder: boolean;
  children?: ITreeNode[];
  isExpanded?: boolean;
  serverRelativeUrl?: string;
  columnInternalName?: string;
  columnValue?: string;
  level: number;
  filterQuery?: string;
  isClicked: boolean;
}

interface IComponentTreeViewState {
  treeData: ITreeNode[];
  loading: boolean;
  error: string;
  allDocumentsCache: any[];
  aplicacaoNormativoListId: string | null;
  iframeUrl: string;
  selectedKey: string | null; // << ADICIONE ESTA LINHA

}

const t = getTranslations();

export default class TreeView extends React.Component<ITreeViewProps, IComponentTreeViewState> {
  constructor(props: ITreeViewProps) {
    super(props);
    this.state = {
      treeData: [],
      loading: true,
      error: "",
      allDocumentsCache: [],
      aplicacaoNormativoListId: null,
      iframeUrl: "",
      selectedKey: null,
    };
  }

  // Verifica se a URL existe (retorna true se status 200-299)
  private async checkUrlExists(url: string): Promise<boolean> {
    try {
      const res = await fetch(url, { method: "HEAD" });
      return res.ok;
    } catch {
      return false;
    }
  }

  private async getDefaultLibraryViewUrl(): Promise<string> {
    const webUrl = this.props.context.pageContext.web.absoluteUrl;
    const libUrl = this.props.selectedLibraryUrl || "";
    if (!libUrl) return "";

    const libAbs = new URL(libUrl, webUrl); // https://<tenant>/sites/.../Normativos
    const libNoTrail = libAbs.href.replace(/\/$/, "");

    // Define página conforme idioma
    const lang = (getUserLanguage() || "pt").toLowerCase();
    const viewPage = lang.startsWith("es") ? "ES.aspx" : "PT.aspx";
    const candidateUrl = `${libNoTrail}/Forms/${viewPage}`;

    // Verifica se a view existe
    const exists = await this.checkUrlExists(candidateUrl);

    if (exists) {
      return candidateUrl;
    }

    // Fallback para raiz da biblioteca (sem .aspx)
    return libNoTrail;
  }



  public async componentDidMount(): Promise<void> {
    injectCssStringOnce(HIDE_SWITCHER_CSS, 'treeview_hide_switcher_css');

    const pageUrl = TreeViewConfigService.getCurrentPageUrl();
    try {
      const cfg = await TreeViewConfigService.loadByPage(pageUrl);

      if (cfg?.PublishedTreeData) {
        const allItems = JSON.parse(cfg.PublishedTreeData);
        this.setState({ allDocumentsCache: allItems });
        this.buildTreeFromData(allItems);

        // const defaultUrl = await this.getDefaultLibraryViewUrl();
        // this.setState({ iframeUrl: defaultUrl });

        return;
      }
    } catch (err) {
      console.warn("Não foi possível ler config/JSON publicado. Seguindo com fluxo online.", err);
    }

    await this.checkAndLoadCache();

    // const defaultUrl = await this.getDefaultLibraryViewUrl();
    // this.setState({ iframeUrl: defaultUrl });
  }


  public async componentDidUpdate(prevProps: ITreeViewProps): Promise<void> {
    if (
      this.props.selectedLibraryUrl !== prevProps.selectedLibraryUrl ||
      this.props.metadataColumn1 !== prevProps.metadataColumn1 ||
      this.props.metadataColumn2 !== prevProps.metadataColumn2 ||
      this.props.metadataColumn3 !== prevProps.metadataColumn3
    ) {
      // Se as propriedades mudaram, limpa o cache e recarrega os dados.
      sessionStorage.removeItem('treeViewCacheData');
      await this.checkAndLoadCache();
    }
  }

  /**
   * Método para verificar e carregar os dados do cache.
   * Se o cache não existir ou estiver incompleto, chama a função para buscar os dados.
   */
  private async checkAndLoadCache(): Promise<void> {
    const { metadataColumn1, metadataColumn2, metadataColumn3 } = this.props;

    const cacheKey = 'treeViewCacheData';
    const cachedData = sessionStorage.getItem(cacheKey);

    if (cachedData) {
      const parsedData = JSON.parse(cachedData);
      const cachedColumns = parsedData.columns;
      const allItems = parsedData.items;

      const currentColumns = [metadataColumn1, metadataColumn2, metadataColumn3].filter(Boolean);
      const cacheIsValid = JSON.stringify(currentColumns.sort()) === JSON.stringify(cachedColumns.sort());

      if (cacheIsValid) {
        this.setState({ allDocumentsCache: allItems });
        this.buildTreeFromData(allItems);


        return;
      }
    }

    // Se o cache não for válido, busca os dados da API
    await this.loadTreeData();
  }

  private buildTreeFromData(allItems: any[]): void {
    const { selectedLibraryUrl, selectedLibraryTitle, metadataColumn1 } = this.props;
    const libraryRootNode: ITreeNode = {
      key: selectedLibraryUrl,
      label: selectedLibraryTitle,
      icon: "Library",
      isFolder: true,
      serverRelativeUrl: selectedLibraryUrl,
      children: [],
      isExpanded: true,
      level: 0,
      filterQuery: "",
      isClicked: false
    };

    let firstLevelNodes: ITreeNode[] = [];
    if (metadataColumn1) {
      firstLevelNodes = this.buildMetadataTreeLevel(1, [], allItems);
    } else {
      firstLevelNodes = this.getDocumentsInThisScope(allItems);
    }

    libraryRootNode.children = firstLevelNodes;
    this.setState({ treeData: [libraryRootNode], loading: false, error: "" });
  }

  private async loadTreeData(): Promise<void> {
    const {
      selectedLibraryUrl,
      selectedLibraryTitle,
      metadataColumn1,
      metadataColumn2,
      metadataColumn3,
      metadataColumnTypes
    } = this.props;

    if (!selectedLibraryUrl) {
      this.setState({
        loading: false,
        error: t.noLibrary
      });
      return;
    }

    this.setState({ loading: true, error: "" });

    try {
      const listInfo = (await pnp.sp.web.lists
        .filter(`RootFolder/ServerRelativeUrl eq '${selectedLibraryUrl}'`)
        .select("Id")
        .get())[0];

      if (!listInfo?.Id) {
        throw new Error(t.error_library_url_not_found);
      }

      if (metadataColumn1 === "aplicacaoNormativo" ||
        metadataColumn2 === "aplicacaoNormativo" ||
        metadataColumn3 === "aplicacaoNormativo") {
        await this.getAplicacaoNormativoListId(listInfo.Id);
      }

      const columnsToProcess = [metadataColumn1, metadataColumn2, metadataColumn3].filter(Boolean);
      const finalSelectColumns = ["ID", "FileRef", "FileLeafRef", "ContentTypeId", "FSObjType"];
      const expandStatements: string[] = [];

      columnsToProcess.forEach(col => {
        if (!col) return;

        let select = col;
        let expand: string | undefined;

        if (col === "aplicacaoNormativo") {
          const lang = getUserLanguage();
          if (lang === "pt") {
            select = `${col}/Id,${col}/DescTipoAplicacaoPT`;
          } else if (lang === "es") {
            select = `${col}/Id,${col}/DescTipoAplicacaoES`;
          } else {
            select = `${col}/Id,${col}/DescTipoAplicacaoPT`;
          }
          expand = col;
        } else {
          const colMeta = metadataColumnTypes?.[col];
          if (colMeta && (colMeta.type === "Lookup" || colMeta.type === "User" || colMeta.type === "ManagedMetadata")) {
            const field = colMeta.lookupField || "Title";
            select = `${col}/Id,${col}/${field}`;
            expand = col;
          } else if ((col.endsWith("0") || col.endsWith("_0")) && !col.includes("/")) {
            select = col.endsWith("_0") ? col.slice(0, -2) : col.slice(0, -1);
            expand = select;
          } else if (col.includes("/")) {
            expand = col.split("/")[0];
          }
        }

        finalSelectColumns.push(select);
        if (expand && !expandStatements.includes(expand)) {
          expandStatements.push(expand);
        }
      });

      const allItems = await pnp.sp.web.lists.getById(listInfo.Id).items
        .select(...finalSelectColumns)
        .expand(...expandStatements)
        .getAll();

      // Salva os dados e as colunas usadas no cache
      sessionStorage.setItem('treeViewCacheData', JSON.stringify({
        items: allItems,
        columns: columnsToProcess
      }));

      this.setState({ allDocumentsCache: allItems });
      this.buildTreeFromData(allItems);



      try {
        const pageUrl = TreeViewConfigService.getCurrentPageUrl();
        const hierarchy = JSON.stringify(
          [metadataColumn1, metadataColumn2, metadataColumn3].filter(Boolean)
        );
        const library = selectedLibraryUrl || "";

        await TreeViewConfigService.upsertPublishedData(
          pageUrl,
          JSON.stringify(allItems),
          library,
          hierarchy
        );
      } catch (e) {
        console.warn("Falha ao publicar JSON na lista TreeViewConfigs", e);
      }

    } catch (error) {
      this.setState({ error: `${t.error_loading_data} ${escape((error as any).message)}`, loading: false, treeData: [], allDocumentsCache: [] });
    }
  }

  /**
   * Obtém o ID da lista referenciada pela coluna "aplicacaoNormativo".
   * @param listId O ID da lista de origem (biblioteca de documentos).
   */
  private async getAplicacaoNormativoListId(listId: string): Promise<void> {
    try {
      const fieldInfo = await pnp.sp.web.lists.getById(listId).fields
        .filter(`InternalName eq 'aplicacaoNormativo'`)
        .select('LookupList')
        .get();

      if (fieldInfo && fieldInfo.length > 0) {
        this.setState({ aplicacaoNormativoListId: fieldInfo[0].LookupList });
      }
    } catch (error) {
      console.error("Erro ao obter o ID da lista 'aplicacaoNormativo':", error);
    }
  }

  private getDocumentsInThisScope = (docs: any[]): ITreeNode[] =>
    docs.filter(doc =>
      doc.FileRef && doc.FileLeafRef &&
      (doc.FSObjType === 0 || (doc.ContentTypeId && !doc.ContentTypeId.startsWith("0x0120")))
    ).map(doc => ({
      key: doc.FileRef,
      label: doc.FileLeafRef,
      icon: this.getFileIcon(doc.FileLeafRef),
      url: doc.FileRef + '?web=1',
      isFolder: false,
      level: 99,
      isClicked: false
    }));

  private buildMetadataTreeLevel = (
    currentLevel: number,
    currentFilters: { column: string; value: string }[],
    docs: any[]
  ): ITreeNode[] => {
    const columns = [this.props.metadataColumn1, this.props.metadataColumn2, this.props.metadataColumn3].filter(Boolean);
    if (currentLevel > columns.length) return this.getDocumentsInThisScope(docs);

    const col = columns[currentLevel - 1];
    if (!col) return [];
    const colStr: string = String(col);

    const unique = new Set<string>();
    docs.forEach(doc => {
      const val = this.getFieldValue(doc, col);
      if (val) unique.add(String(val));
    });

    return Array.from(unique).sort().map(value => ({
      key: `${colStr}-${value}-${currentLevel}-${this.buildFilterQueryForItems([...currentFilters, { column: colStr, value }])}`,
      label: this.getLabelWithOptionalId(colStr, value, docs),
      icon: "Tag",
      isFolder: true,
      level: currentLevel,
      columnInternalName: colStr,
      columnValue: value,
      children: [],
      isExpanded: false,
      isClicked: false,
      filterQuery: this.buildFilterQueryForItems([...currentFilters, { column: colStr, value }])
    }));

  };

  private getFieldValue = (item: any, name: string): any => {
    if (!item || !name) return "";

    let val = item[name];
    if (val !== undefined) {
      if (typeof val === "object" && val !== null) {
        if (Array.isArray(val)) {
          return val.map(v =>
            v?.Title ?? v?.Label ?? v?.LookupValue ?? v?.Sigla ?? String(v)
          ).join("; ");
        }
        return val.Title ?? val.Label ?? val.LookupValue ?? val.Sigla ?? val.DescTipoAplicacaoPT ?? val.DescTipoAplicacaoES ?? "";
      }

      if (typeof val === "string" && val.includes(";#")) {
        const parts = val.split(";#");
        return parts[parts.length - 1] ?? "";
      }

      return val;
    }

    if (name.includes("/")) {
      const [base, prop] = name.split("/");
      return item[base]?.[prop] ??
        item[base]?.Title ??
        item[base]?.Label ??
        item[base]?.LookupValue ??
        item[base]?.Sigla ??
        "";
    }

    if (name === "ID") {
      return item[name];
    }
    if (name.endsWith("Id")) {
      const base = name.slice(0, -2);
      return item[base]?.Title ?? item[name] ?? "";
    }


    if (item.ListItemAllFields?.[name] !== undefined) {
      const li = item.ListItemAllFields[name];
      if (typeof li === "object" && li !== null) {
        if (Array.isArray(li)) {
          return li.map(v => v?.Title ?? v?.Label ?? v?.LookupValue ?? String(v)).join("; ");
        }
        return li.Title ?? li.Label ?? li.LookupValue ?? "";
      }

      if (typeof li === "string" && li.includes(";#")) {
        const parts = li.split(";#");
        return parts[parts.length - 1] ?? "";
      }

      return li ?? "";
    }

    return "";
  }

  private buildFilterQueryForItems = (filters: { column: string; value: string }[]): string =>
    filters.map(f => `${f.column} eq '${f.value.replace(/'/g, "''")}'`).join(" and ");

  private getColumnForLevel = (level: number): string | undefined => {
    switch (level) {
      case 1: return this.props.metadataColumn1;
      case 2: return this.props.metadataColumn2;
      case 3: return this.props.metadataColumn3;
      default: return undefined;
    }
  }

  private async handleExpandClick(node: ITreeNode): Promise<void> {
    if (!node.isFolder) {
      return;
    }

    this.setState(prev => ({
      treeData: this.toggleNodeExpansion(prev.treeData, node.key)
    }), async () => {
      const updated = this.findNodeInTree(this.state.treeData, node.key);
      if (updated && updated.isExpanded && updated.children?.length === 0) {
        this.setState({ loading: true });

        const nextLevel = updated.level + 1;
        const column = this.getColumnForLevel(nextLevel);

        const filters = updated.filterQuery?.split(" and ").map(f => {
          const [col, val] = f.split(" eq ");
          return { column: col, value: val.replace(/'/g, "") };
        }) ?? [];

        const scopedDocs = this.state.allDocumentsCache.filter(doc =>
          filters.every(f => String(this.getFieldValue(doc, f.column)) === f.value)
        );

        const children = column
          ? this.buildMetadataTreeLevel(nextLevel, filters, scopedDocs)
          : this.getDocumentsInThisScope(scopedDocs);

        this.setState(prev => ({
          treeData: this.addChildrenToNode(prev.treeData, node.key, children),
          loading: false
        }));
      }
    });
  }


  private async handleNodeClick(node: ITreeNode): Promise<void> {

    if (!node.isFolder) {
      if (node.url) window.open(node.url, '_blank');
      return;
    }

    if (node.level === 0) {
      this.setState({ iframeUrl: "", selectedKey: null });
      return;
    }

    const iframeUrl = await this.buildIframeUrl(node);
    if (iframeUrl) {
      this.setState({ iframeUrl, selectedKey: node.key });
    } else {
      this.setState({ selectedKey: node.key });
    }
  }

  private async buildIframeUrl(node: ITreeNode): Promise<string> {
    const webUrl = this.props.context.pageContext.web.absoluteUrl;
    const libUrl = this.props.selectedLibraryUrl || "";
    if (!libUrl) return "";

    const libAbs = new URL(libUrl, webUrl);
    const libNoTrail = libAbs.href.replace(/\/$/, "");

    // 1. Constrói a URL com a view padrão e os filtros
    const lang = (getUserLanguage() || "pt").toLowerCase();
    const viewPage = lang.startsWith("es") ? "ES.aspx" : "PT.aspx";
    const baseViewWithAspx = `${libNoTrail}/Forms/${viewPage}`;

    const nodePath = this.findNodePath(this.state.treeData, node.key);
    const filterParams: string[] = [];
    let filterCount = 1;

    if (nodePath) {
      for (let i = 1; i < nodePath.length; i++) {
        const currentPathNode = nodePath[i];
        if (currentPathNode.columnInternalName && currentPathNode.columnValue) {
          let filterField = currentPathNode.columnInternalName;
          let filterValue = currentPathNode.columnValue;

          if (currentPathNode.columnInternalName === "aplicacaoNormativo") {
            const aplicacaoNormativoId = this.state.allDocumentsCache
              .find(doc => this.getFieldValue(doc, currentPathNode.columnInternalName) === currentPathNode.columnValue)
              ?.aplicacaoNormativo?.Id;

            if (aplicacaoNormativoId) {
              filterValue = aplicacaoNormativoId;
            }
          }
          filterParams.push(
            `FilterField${filterCount}=${encodeURIComponent(filterField)}`,
            `FilterValue${filterCount}=${encodeURIComponent(filterValue)}`,
            `FilterType${filterCount}=Lookup`
          );
          filterCount++;
        }
      }
    }

    const filtersQuery = filterParams.length > 0 ? `?${filterParams.join("&")}` : "";
    const candidateUrl = `${baseViewWithAspx}${filtersQuery}`;

    // 2. Verifica se a URL com a view e os filtros existe.
    const exists = await this.checkUrlExists(candidateUrl);

    if (exists) {
      console.log("URL do Iframe com filtros concatenados:", candidateUrl);
      return candidateUrl;
    } else {
      // 3. Fallback: se a URL com a view for 404, retorna a URL da biblioteca.
      const fallbackUrl = libNoTrail;
      console.log("URL com .aspx retornou 404. Usando fallback:", fallbackUrl);
      return fallbackUrl;
    }
  }


  // Método auxiliar para encontrar a trilha raiz dos nós da hierarquia, fundamental para fazer a URL do iframe retornar os filtros concatenados.
  private findNodePath = (nodes: ITreeNode[], key: string, path: ITreeNode[] = []): ITreeNode[] | undefined => {
    for (const n of nodes) {
      const newPath = [...path, n];
      if (n.key === key) {
        return newPath;
      }
      if (n.children && n.children.length > 0) {
        const foundPath = this.findNodePath(n.children, key, newPath);
        if (foundPath) {
          return foundPath;
        }
      }
    }
    return undefined;
  }

  private toggleNodeExpansion = (nodes: ITreeNode[], key: string): ITreeNode[] =>
    nodes.map(n => ({
      ...n,
      children: n.children ? this.toggleNodeExpansion(n.children, key) : n.children,
      isExpanded: n.key === key ? !n.isExpanded : n.isExpanded
    }));

  private addChildrenToNode = (nodes: ITreeNode[], key: string, children: ITreeNode[]): ITreeNode[] =>
    nodes.map(n => ({
      ...n,
      children: n.key === key ? children : n.children ? this.addChildrenToNode(n.children, key, children) : n.children
    }));

  private findNodeInTree = (nodes: ITreeNode[], key: string): ITreeNode | undefined => {
    for (const n of nodes) {
      if (n.key === key) return n;
      const found = this.findNodeInTree(n.children ?? [], key);
      if (found) return found;
    }
    return undefined;
  }

  private getFileIcon = (name: string): string => {
    const ext = name.split('.').pop()?.toLowerCase();
    switch (ext) {
      case 'doc': case 'docx': return 'WordDocument';
      case 'xls': case 'xlsx': return 'ExcelDocument';
      case 'ppt': case 'pptx': return 'PowerPointDocument';
      case 'pdf': return 'PDF';
      case 'jpg': case 'jpeg': case 'png': case 'gif': return 'Photo2';
      case 'zip': case 'txt': return 'TextDocument';
      default: return 'Document';
    }
  }

  private getFriendlyColumnValue = (val: string, name: string): string => {
    if (name.toLowerCase().includes("date")) {
      try {
        return new Date(val).toLocaleDateString();
      } catch { return val; }
    }
    return val;
  }

  /**
 * Formata o ID com zero à esquerda para no mínimo 2 dígitos.
 * Ex.: "1" -> "01", "10" -> "10".
 */
  private formatId2Digits(idLike: string | number | null | undefined): string {
    if (idLike === null || idLike === undefined) return "";
    const n = parseInt(String(idLike), 10);
    if (isNaN(n)) return String(idLike);
    return String(n).padStart(2, "0");
  }


  /**
 * Tenta extrair um ID de um valor bruto do campo (objeto, array ou string no formato SharePoint "12;#Rótulo;#34;#Outro").
 */
  private tryExtractIdFromRaw(raw: any, value: string): string | null {
    if (!raw) return null;

    // Caso 1: Objeto único já expandido (Lookup/User/ManagedMetadata)
    if (typeof raw === "object" && !Array.isArray(raw)) {
      // Alguns tipos podem expor Id ou WssId (taxonomia)
      return raw.Id ?? raw.WssId ?? null;
    }

    // Caso 2: Array (multi-lookup/managed metadata multi)
    if (Array.isArray(raw)) {
      // Procura o item do array cujo "Title/Label/LookupValue/Sigla/Desc*" bate com o 'value'
      const hit = raw.find(v => {
        const display =
          v?.Title ?? v?.Label ?? v?.LookupValue ?? v?.Sigla ?? v?.DescTipoAplicacaoPT ?? v?.DescTipoAplicacaoES ?? String(v);
        return String(display) === String(value);
      });
      if (hit) {
        return hit.Id ?? hit.WssId ?? null;
      }
      return null;
    }

    // Caso 3: String no formato "12;#Rótulo" (ou com vários pares)
    if (typeof raw === "string" && raw.includes(";#")) {
      // Divide em pares [id, label, id, label, ...]
      const parts = raw.split(";#");
      // Ex.: ["12", "Rótulo", "34", "Outro"]
      for (let i = 0; i < parts.length - 1; i += 2) {
        const maybeId = parts[i];
        const maybeLabel = parts[i + 1];
        if (String(maybeLabel) === String(value)) {
          return maybeId || null;
        }
      }
      // fallback simples para string "id;#label" (último par)
      if (parts.length >= 2 && parts[parts.length - 1] === value) {
        return parts[parts.length - 2] || null;
      }
    }

    return null;
  }


  private getIdForColumnValue(col: string, value: string, docs: any[]): string | null {
    if (!col || !value) return null;

    if (col !== "aplicacaoNormativo") return null;

    const match = docs.find(doc => String(this.getFieldValue(doc, col)) === String(value));
    if (!match) return null;

    if (match[col]?.Id) {
      return String(match[col].Id);
    }

    const raw = match[col];
    const idFromRaw = this.tryExtractIdFromRaw(raw, value);
    if (idFromRaw) return String(idFromRaw);

    if (col.includes("/")) {
      const base = col.split("/")[0];
      const idFromBase = this.tryExtractIdFromRaw(match[base], value);
      if (idFromBase) return String(idFromBase);
      if (match[base]?.Id) return String(match[base].Id);
    }

    if (col.endsWith("Id") && String(col) !== "ID") {
      const idVal = match[col];
      if (idVal !== undefined && idVal !== null && String(this.getFieldValue(match, col.slice(0, -2))) === String(value)) {
        return String(idVal);
      }
    }

    const li = match.ListItemAllFields?.[col];
    if (li) {
      const idFromLi = this.tryExtractIdFromRaw(li, value);
      if (idFromLi) return String(idFromLi);
      if (li?.Id) return String(li.Id);
    }

    return null;
  }

  /**
   * Monta o rótulo que será exibido na árvore: "ID - ValorAmigável" se houver ID, senão "ValorAmigável".
   */
  private getLabelWithOptionalId(col: string, rawValue: string, docsInScope: any[]): string {
    const friendly = this.getFriendlyColumnValue(rawValue, col);

    if (col === "aplicacaoNormativo") {
      const id = this.getIdForColumnValue(col, friendly, docsInScope);
      if (id) {
        const padded = this.formatId2Digits(id);
        return `${padded} - ${friendly}`;
      }
    }

    return friendly;
  }



  public render(): React.ReactElement<ITreeViewProps> {
    const { loading, error, treeData, iframeUrl, selectedKey } = this.state;

    const renderTreeNodes = (nodes: ITreeNode[]) => (
      <ul className={styles.treeList}>
        {nodes.map(node => (
          <li key={node.key} className={styles.treeNode}>
            <div className={styles.nodeContent}>
              {node.isFolder && (
                <span className={styles.expanderIcon} onClick={() => this.handleExpandClick(node)}>
                  {node.isExpanded ? "▼" : "►"}
                </span>
              )}
              <Icon iconName={node.icon} style={{ marginRight: 5 }} />
              <span
                onClick={() => this.handleNodeClick(node)}
                className={node.key === selectedKey ? styles.selectedNodeLabel : undefined}
              >
                {escape(node.label)}
              </span>
            </div>
            {node.isFolder && node.isExpanded && (
              <div className={styles.childrenContainer}>
                {node.children?.length
                  ? renderTreeNodes(node.children)
                  : loading && <div className={styles.loadingIndicator}>Carregando...</div>}
              </div>
            )}
          </li>
        ))}
      </ul>
    );

    return (
      <section className={`${styles.treeViewContainer} ${this.props.hasTeamsContext}`}>
        <SplitterLayout
          percentage
          primaryIndex={1}
          secondaryInitialSize={20}
          primaryMinSize={40}
          secondaryMinSize={10}
        >
          {/* Painel esquerdo: Árvore */}
          <div className={styles.treeView}>
            {/* <p>{t.welcome.replace('{user}', this.props.userDisplayName)}</p> */}
            <div className={styles.treeContainer}>
              {loading && treeData.length === 0 && <p>{t.loading}</p>}
              {error && <p style={{ color: 'red' }}>{error}</p>}
              {!loading && !error && treeData.length === 0 && (
                <p>{!this.props.selectedLibraryUrl
                  ? t.noLibrary
                  : (!this.props.metadataColumn1 && !this.props.metadataColumn2 && !this.props.metadataColumn3)
                    ? t.noMetadata
                    : t.noDocuments}</p>
              )}
              {!loading && !error && treeData.length > 0 && renderTreeNodes(treeData)}
            </div>
          </div>

          {/* Painel direito: IframePreview */}
          <div className={styles.iframeContainer}>
            <IframePreview
              url={iframeUrl}
              listTitle={this.props.selectedLibraryTitle}
              emptyMessage={t.select_item_to_show_normativos}
            />
          </div>
        </SplitterLayout>
      </section>
    );
  }
}