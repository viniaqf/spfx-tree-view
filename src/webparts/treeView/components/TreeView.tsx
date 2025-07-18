import * as React from 'react';
import styles from './TreeView.module.scss';
import { ITreeViewProps } from './ITreeViewProps';
import pnp from "sp-pnp-js";
import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from 'office-ui-fabric-react';

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
}

interface IComponentTreeViewState {
  treeData: ITreeNode[];
  loading: boolean;
  error: string;
  allDocumentsCache: any[];
}

export default class TreeView extends React.Component<ITreeViewProps, IComponentTreeViewState> {
  constructor(props: ITreeViewProps) {
    super(props);
    this.state = {
      treeData: [],
      loading: true,
      error: "",
      allDocumentsCache: []
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadTreeData();
  }

  public async componentDidUpdate(prevProps: ITreeViewProps): Promise<void> {
    if (
      this.props.selectedLibraryUrl !== prevProps.selectedLibraryUrl ||
      this.props.metadataColumn1 !== prevProps.metadataColumn1 ||
      this.props.metadataColumn2 !== prevProps.metadataColumn2 ||
      this.props.metadataColumn3 !== prevProps.metadataColumn3
    ) {
      await this.loadTreeData();
    }
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
        error: "Por favor, selecione uma biblioteca de documentos nas configurações da Web Part."
      });
      return;
    }

    this.setState({ loading: true, error: "" });

    try {
      const libraryRootNode: ITreeNode = {
        key: selectedLibraryUrl,
        label: selectedLibraryTitle || "Biblioteca Selecionada",
        icon: "Library",
        isFolder: true,
        serverRelativeUrl: selectedLibraryUrl,
        children: [],
        isExpanded: true,
        level: 0,
        filterQuery: ""
      };

      const listInfo = (await pnp.sp.web.lists
        .filter(`RootFolder/ServerRelativeUrl eq '${selectedLibraryUrl}'`)
        .select("Id")
        .get())[0];

      if (!listInfo?.Id) {
        throw new Error("Não foi possível encontrar a lista para a URL da biblioteca fornecida.");
      }

      const columnsToProcess = [metadataColumn1, metadataColumn2, metadataColumn3].filter(Boolean);
      const finalSelectColumns = ["ID", "FileRef", "FileLeafRef", "ContentTypeId", "FSObjType"];
      const expandStatements: string[] = [];

      columnsToProcess.forEach(col => {
        if (!col) return;

        let select = col;
        let expand: string | undefined;

        if (col === "aplicacaoNormativo") {
          // Caso especial — somente para aplicaçãoNormativo, pois ele referencia um ID, que não reconhece o valor de aplicacaoNormativo. Portanto, precisamos especificar a utilização do campo de referência correto para esse caso: "/DescTipoAplicacaoPT".  
          select = `${col}/DescTipoAplicacaoPT`;
          expand = col;
        } else {
          const colMeta = metadataColumnTypes?.[col];
          if (colMeta && (colMeta.type === "Lookup" || colMeta.type === "User" || colMeta.type === "ManagedMetadata")) {
            const field = colMeta.lookupField || "Title";
            select = `${col}/${field}`;
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

      this.setState({ allDocumentsCache: allItems });

      let firstLevelNodes: ITreeNode[] = [];
      if (metadataColumn1) {
        firstLevelNodes = this.buildMetadataTreeLevel(1, [], allItems);
      } else {
        firstLevelNodes = this.getDocumentsInThisScope(allItems);
      }

      libraryRootNode.children = firstLevelNodes;
      this.setState({ treeData: [libraryRootNode], loading: false });

    } catch (error) {
      console.error("Erro ao carregar a árvore de metadados:", error);
      this.setState({ error: `Erro ao carregar dados: ${escape(error.message)}`, loading: false, treeData: [], allDocumentsCache: [] });
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
      url: doc.FileRef + `?web=1`,
      isFolder: false,
      level: 99
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

    const unique = new Set<string>();
    docs.forEach(doc => {
      const val = this.getFieldValue(doc, col);
      if (val) unique.add(String(val));
    });

    return Array.from(unique).sort().map(value => ({
      key: `${col}-${value}-${currentLevel}-${this.buildFilterQueryForItems([...currentFilters, { column: col, value }])}`,
      label: this.getFriendlyColumnValue(value, col),
      icon: "Tag",
      isFolder: true,
      level: currentLevel,
      columnInternalName: col,
      columnValue: value,
      children: [],
      isExpanded: false,
      filterQuery: this.buildFilterQueryForItems([...currentFilters, { column: col, value }])
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
        return val.Title ?? val.Label ?? val.LookupValue ?? val.Sigla ?? val.DescTipoAplicacaoPT ?? "";
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

    if (name.endsWith("Id") && name !== "ID") {
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

  private async handleNodeClick(node: ITreeNode): Promise<void> {
    if (!node.isFolder) {
      window.open(node.url, '_blank');
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

  public render(): React.ReactElement<ITreeViewProps> {
    const { loading, error, treeData } = this.state;

    const renderTreeNodes = (nodes: ITreeNode[]) => (
      <ul className={styles.treeList}>
        {nodes.map(node => (
          <li key={node.key} className={styles.treeNode}>
            <div className={styles.nodeContent} onClick={() => this.handleNodeClick(node)}>
              {node.isFolder && <span className={styles.expanderIcon}>{node.isExpanded ? "▼" : "►"}</span>}
              <Icon iconName={node.icon} style={{ marginRight: 5 }} />
              <span>{escape(node.label)}</span>
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
      <section className={`${styles.treeView} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Visualizador de Documentos por Metadados</h2>
          <div>Web part property value: <strong>{escape(this.props.description)}</strong></div>
        </div>
        <div className={styles.treeContainer}>
          {loading && treeData.length === 0 && <p>Carregando...</p>}
          {error && <p style={{ color: 'red' }}>{error}</p>}
          {!loading && !error && treeData.length === 0 && (
            <p>{!this.props.selectedLibraryUrl
              ? "Por favor, abra as configurações da Web Part e selecione uma biblioteca de documentos."
              : (!this.props.metadataColumn1 && !this.props.metadataColumn2 && !this.props.metadataColumn3)
                ? "Nenhuma coluna de metadados selecionada. Exibindo documentos da raiz da biblioteca."
                : "A biblioteca selecionada não contém documentos com os metadados especificados."}</p>
          )}
          {!loading && !error && treeData.length > 0 && renderTreeNodes(treeData)}
        </div>
      </section>
    );
  }
}
