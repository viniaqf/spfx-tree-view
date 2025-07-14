// src/webparts/treeView/components/TreeView.tsx

import * as React from 'react';
import styles from './TreeView.module.scss';
import { ITreeViewProps } from './ITreeViewProps';

import pnp from "sp-pnp-js";

import { escape } from '@microsoft/sp-lodash-subset';


// --- Interfaces para a estrutura da árvore (nós) ---
interface ITreeNode {
  key: string;
  label: string;
  icon?: string;
  url?: string;
  isFolder: boolean; // True para nós de metadados/biblioteca, False para documentos
  children?: ITreeNode[];
  isExpanded?: boolean;
  serverRelativeUrl?: string;
  columnInternalName?: string;
  columnValue?: string;
  level: number;
  filterQuery?: string;
}

// Interface para o estado interno do componente TreeView
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
    if (this.props.selectedLibraryUrl !== prevProps.selectedLibraryUrl ||
        this.props.metadataColumn1 !== prevProps.metadataColumn1 ||
        this.props.metadataColumn2 !== prevProps.metadataColumn2 ||
        this.props.metadataColumn3 !== prevProps.metadataColumn3) {
      await this.loadTreeData();
    }
  }

  private async loadTreeData(): Promise<void> {
    const { selectedLibraryUrl, selectedLibraryTitle, metadataColumn1, metadataColumn2, metadataColumn3, metadataColumnTypes } = this.props;

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

      const listInfo = (await pnp.sp.web.lists.filter(`RootFolder/ServerRelativeUrl eq '${selectedLibraryUrl}'`).select("Id").get())[0];
      if (!listInfo || !listInfo.Id) {
          throw new Error("Não foi possível encontrar a lista para a URL da biblioteca fornecida.");
      }

      const columnsToProcess = [
        metadataColumn1,
        metadataColumn2,
        metadataColumn3
      ].filter(Boolean);

      const finalSelectColumns = ["ID", "FileRef", "FileLeafRef", "ContentTypeId", "FSObjType"]; // Mantenho FSObjType aqui para filtro cliente
      const expandStatements: string[] = [];

      columnsToProcess.forEach(colInternalName => {
        if (!colInternalName) return;

        let selectString = colInternalName; 
        let expandPartForThisCol: string | undefined = undefined; 

        // --- TRATAMENTO ESPECÍFICO PARA 'siglaDoTipoDoNormativo' ---
        if (colInternalName === "siglaDoTipoDoNormativo") {
            selectString = `${colInternalName}/Sigla`;
            expandPartForThisCol = colInternalName; 
        }
        // Heurística para campos que terminam em '0' ou '_0' (Managed Metadata Text Field)
        else if ((colInternalName.endsWith("0") || colInternalName.endsWith("_0")) && !colInternalName.includes("/")) {
            selectString = colInternalName.substring(0, colInternalName.length - 1); 
            if (colInternalName.endsWith("_0")) {
                selectString = colInternalName.substring(0, colInternalName.length - 2);
            }
            expandPartForThisCol = selectString; 
        }
        // Heurística genérica para outros campos Lookup/Pessoa/Managed Metadata (que não foram tratados acima)
        else if (colInternalName.includes("/") || 
                 colInternalName.endsWith("Id") || 
                 colInternalName.toLowerCase().includes("lookup") ||
                 colInternalName.toLowerCase().includes("user") ||
                 colInternalName.toLowerCase().includes("person") ||
                 colInternalName.toLowerCase().includes("editor") || 
                 colInternalName.toLowerCase().includes("author") || 
                 colInternalName.toLowerCase().includes("modifiedby") || 
                 colInternalName.toLowerCase().includes("createdby") || 
                 colInternalName.toLowerCase().includes("managedmetadata") ||
                 colInternalName === "Area_x0020_Gestora" || 
                 colInternalName === "TituloPT" 
                ) {
            if (!colInternalName.includes("/") && !colInternalName.endsWith("Id")) {
                selectString = `${colInternalName}/Title`; 
            } else {
                selectString = colInternalName;
            }
            
            const baseColNameForExpand = colInternalName.split('/')[0];
            expandPartForThisCol = baseColNameForExpand;
        } 
        // Campos simples
        else {
            selectString = colInternalName; 
        }
        
        finalSelectColumns.push(selectString);

        if (expandPartForThisCol && !expandStatements.includes(expandPartForThisCol)) {
            expandStatements.push(expandPartForThisCol);
        }
      });

      const allItemsInLibrary = await pnp.sp.web.lists.getById(listInfo.Id).items
                                                   .select(...finalSelectColumns)
                                                   // REMOVIDO AQUI: .filter("FileSystemObjectType eq 0")
                                                   .expand(...expandStatements.filter(Boolean))
                                                   .getAll(); 

      this.setState({ allDocumentsCache: allItemsInLibrary });

      let firstLevelNodes: ITreeNode[] = [];
      if (metadataColumn1) {
        firstLevelNodes = this.buildMetadataTreeLevel(
          1,
          [],
          allItemsInLibrary
        );
      } else {
        firstLevelNodes = this.getDocumentsInThisScope(allItemsInLibrary);
      }
      libraryRootNode.children = firstLevelNodes;

      this.setState({ treeData: [libraryRootNode], loading: false });

    } catch (error) {
      console.error("Erro ao carregar a árvore de metadados:", error);
      this.setState({ error: `Erro ao carregar dados: ${escape(error.message)}`, loading: false, treeData: [], allDocumentsCache: [] });
    }
  }

  // --- getDocumentsInThisScope agora filtra SOMENTE no CLIENTE ---
  private getDocumentsInThisScope = (documentsInScope: any[]): ITreeNode[] => {
    return documentsInScope
      // Filtra itens que são documentos:
      // 1. Possuem FileRef E FileLeafRef (são arquivos, não apenas itens ou pastas sem arquivo)
      // 2. Têm FSObjType como 0 (padrão de arquivo CAML) OU ContentTypeId que NÃO começa com "0x0120" (padrão de pasta)
      .filter(doc => 
        doc.FileRef && doc.FileLeafRef && 
        (doc.FSObjType === 0 || (doc.ContentTypeId && !doc.ContentTypeId.startsWith("0x0120")))
      ) 
      .map(doc => ({
        key: doc.FileRef,
        label: doc.FileLeafRef,
        icon: this.getFileIcon(doc.FileLeafRef), 
        url: doc.FileRef,
        isFolder: false, 
        level: 99
      }));
  }


  private buildMetadataTreeLevel = (
    currentLevel: number,
    currentFilters: { column: string; value: string; }[],
    documentsInScope: any[]
  ): ITreeNode[] => {
    const metadataColumns = [
      this.props.metadataColumn1,
      this.props.metadataColumn2,
      this.props.metadataColumn3
    ].filter(Boolean);

    if (currentLevel > metadataColumns.length) {
      return this.getDocumentsInThisScope(documentsInScope);
    }

    const currentColumnInternalName = metadataColumns[currentLevel - 1];
    if (!currentColumnInternalName) {
        return [];
    }

    const uniqueValues = new Set<string>();
    documentsInScope.forEach(doc => {
      const fieldValue = this.getFieldValue(doc, currentColumnInternalName);
      if (fieldValue !== undefined && fieldValue !== null && fieldValue !== "") {
        uniqueValues.add(String(fieldValue)); 
      }
    });

    return Array.from(uniqueValues).sort().map(value => ({
      key: `${currentColumnInternalName}-${value}`,
      label: this.getFriendlyColumnValue(value, currentColumnInternalName),
      icon: "Tag",
      isFolder: true,
      level: currentLevel,
      columnInternalName: currentColumnInternalName,
      columnValue: value, 
      children: [],
      isExpanded: false,
      filterQuery: this.buildFilterQueryForItems([...currentFilters, { column: currentColumnInternalName, value: value }])
    }));
  }

  private getFieldValue = (item: any, internalName: string): any => {
    if (!item || !internalName) { return undefined; }

    const directValue = item[internalName];
    if (directValue !== undefined) {
        if (typeof directValue === 'object' && directValue !== null) {
            if (directValue.Title !== undefined) { return directValue.Title || ""; }
            if (directValue.Label !== undefined) { return directValue.Label || ""; }
            if (directValue.LookupValue !== undefined) { return directValue.LookupValue || ""; }
            // Adicionado: Lógica para campo 'Sigla' (ou outros com nomes específicos)
            if (directValue.Sigla !== undefined) { return directValue.Sigla || ""; }
            
            if (Array.isArray(directValue)) {
                return directValue.map(val => {
                    if (val && val.Title !== undefined) return val.Title || "";
                    if (val && val.Label !== undefined) return val.Label || "";
                    if (val && val.LookupValue !== undefined) return val.LookupValue || "";
                    if (val && val.Sigla !== undefined) return val.Sigla || ""; // Para array de objetos
                    return String(val);
                }).join('; ');
            }
        }
        if (typeof directValue === 'string' && directValue.includes(';#')) {
            const parts = directValue.split(';#');
            if (parts.length > 1) { return parts[parts.length - 1] || ""; }
        }
        return directValue;
    }

    if (internalName.includes('/')) {
        const [complexFieldName, complexProp] = internalName.split('/');
        if (item[complexFieldName] && item[complexFieldName][complexProp] !== undefined) {
            return item[complexFieldName][complexProp] || "";
        }
        if (item[complexFieldName] !== undefined && typeof item[complexFieldName] === 'object') {
            const baseValue = item[complexFieldName];
            if (baseValue && baseValue.Title !== undefined) { return baseValue.Title || ""; }
            if (baseValue && baseValue.Label !== undefined) { return baseValue.Label || ""; }
            if (baseValue && baseValue.LookupValue !== undefined) { return baseValue.LookupValue || ""; }
            if (baseValue && baseValue.Sigla !== undefined) { return baseValue.Sigla || ""; } // <-- Adicionado
        }
    }
    
    if (internalName.endsWith("Id") && internalName.length > 2 && internalName !== "ID") {
        const baseFieldName = internalName.substring(0, internalName.length - 2); 
        if (item[baseFieldName] && typeof item[baseFieldName] === 'object' && item[baseFieldName].Title !== undefined) {
            return item[baseFieldName].Title || "";
        }
        if (item[internalName] !== undefined) { return item[internalName]; } 
    }

    if (item.ListItemAllFields && item.ListItemAllFields[internalName] !== undefined) {
        const liValue = item.ListItemAllFields[internalName];
        if (typeof liValue === 'object' && liValue !== null) {
            if (liValue.Title !== undefined) { return liValue.Title || ""; }
            if (liValue.Label !== undefined) { return liValue.Label || ""; }
            if (liValue.LookupValue !== undefined) { return liValue.LookupValue || ""; }
            if (Array.isArray(liValue)) {
                return liValue.map(val => {
                    if (val && val.Title !== undefined) return val.Title || "";
                    if (val && val.Label !== undefined) return val.Label || "";
                    if (val && val.LookupValue !== undefined) return val.LookupValue || "";
                    return String(val);
                }).join('; ');
            }
        }
        if (typeof liValue === 'string' && liValue.includes(';#')) {
            const parts = liValue.split(';#');
            if (parts.length > 1) { return parts[parts.length - 1] || ""; }
        }
        return liValue || "";
    }

    return "";
  }

  private getFileIcon = (fileName: string): string => {
    const extension = fileName.split('.').pop()?.toLowerCase();
    switch (extension) {
      case 'docx': case 'doc': return 'WordDocument';
      case 'xlsx': case 'xls': return 'ExcelDocument';
      case 'pptx': case 'ppt': return 'PowerPointDocument';
      case 'pdf': return 'PDF';
      case 'jpg': case 'jpeg': case 'png': case 'gif': return 'Photo2';
      case 'zip': case 'txt': return 'TextDocument';
      default: return 'Document';
    }
  }

  private getFriendlyColumnValue = (value: string, internalName: string): string => {
    if (internalName.toLowerCase().includes("date")) {
      try {
        return new Date(value).toLocaleDateString();
      } catch (e) { /* ignore */ }
    }
    return value;
  }

  private getColumnForLevel = (level: number): string | undefined => {
    switch (level) {
      case 1: return this.props.metadataColumn1;
      case 2: return this.props.metadataColumn2;
      case 3: return this.props.metadataColumn3;
      default: return undefined;
    }
  }

  private buildFilterQueryForItems = (filters: { column: string; value: string; }[]): string => {
    if (!filters || filters.length === 0) {
        return "";
    }
    const filterParts: string[] = [];
    filters.forEach(f => {
        let value = f.value;
        if (typeof value === 'string' && value.includes("'")) {
            value = value.replace(/'/g, "''");
        }
        
        filterParts.push(`${f.column} eq '${value}'`);
    });
    return filterParts.join(' and ');
  }

  private handleNodeClick = async (node: ITreeNode): Promise<void> => {
    if (!node.isFolder) {
      if (node.url) {
        window.open(node.url, '_blank');
      }
      return;
    }

    this.setState(prevState => {
      const newTreeData = this.toggleNodeExpansion(prevState.treeData, node.key);
      return { treeData: newTreeData };
    }, async () => {
      const updatedNode = this.findNodeInTree(this.state.treeData, node.key);
      if (updatedNode && updatedNode.isExpanded && updatedNode.children && updatedNode.children.length === 0) {
        this.setState({ loading: true });

        const nextLevel = updatedNode.level + 1;
        const nextColumnInternalName = this.getColumnForLevel(nextLevel);
        
        let children: ITreeNode[] = []; 

        const documentsInScopeForChildren = this.state.allDocumentsCache.filter(doc => {
            const docFilterQuery = updatedNode.filterQuery;
            if (!docFilterQuery) return true;

            const filters = docFilterQuery.split(' and ').map(part => {
                const eqIndex = part.indexOf(' eq ');
                if (eqIndex > -1) {
                    const col = part.substring(0, eqIndex);
                    const val = part.substring(eqIndex + 4).replace(/'/g, '');
                    return { column: col, value: val };
                }
                return { column: '', value: '' };
            }).filter(f => f.column);

            return filters.every(f => {
                const fieldValue = this.getFieldValue(doc, f.column);
                if (typeof fieldValue === 'string' && fieldValue.includes(';#')) {
                    return fieldValue.split(';#').some(part => part === f.value);
                }
                return String(fieldValue) === String(f.value);
            });
        });


        if (nextColumnInternalName) {
          children = this.buildMetadataTreeLevel(
            nextLevel,
            updatedNode.filterQuery ? updatedNode.filterQuery.split(' and ').map(p => {
                const [col, val] = p.split(' eq ');
                return { column: col, value: val.replace(/'/g, '') };
            }) : [],
            documentsInScopeForChildren
          );
        } else {
          children = this.getDocumentsInThisScope(documentsInScopeForChildren);
        }

        this.setState(prevState => {
          const treeDataWithChildren = this.addChildrenToNode(prevState.treeData, updatedNode.key, children);
          return { treeData: treeDataWithChildren, loading: false };
        });
      }
    });
  };

  private toggleNodeExpansion = (nodes: ITreeNode[], keyToToggle: string): ITreeNode[] => {
    return nodes.map(node => {
      if (node.key === keyToToggle) {
        return { ...node, isExpanded: !node.isExpanded };
      }
      if (node.children) {
        return { ...node, children: this.toggleNodeExpansion(node.children, keyToToggle) };
      }
      return node;
    });
  };

  private addChildrenToNode = (nodes: ITreeNode[], parentKey: string, children: ITreeNode[]): ITreeNode[] => {
    return nodes.map(node => {
      if (node.key === parentKey) {
        return { ...node, children: children };
      }
      if (node.children) {
        return { ...node, children: this.addChildrenToNode(node.children, parentKey, children) };
      }
      return node;
    });
  };

  private findNodeInTree = (nodes: ITreeNode[], keyToFind: string): ITreeNode | undefined => {
    for (const node of nodes) {
      if (node.key === keyToFind) {
        return node;
      }
      if (node.children) {
        const found = this.findNodeInTree(node.children, keyToFind);
        if (found) { return found; }
      }
    }
    return undefined;
  };

  public render(): React.ReactElement<ITreeViewProps> {
    const { loading, error, treeData } = this.state;
    const { selectedLibraryUrl, metadataColumn1, metadataColumn2, metadataColumn3 } = this.props;

    const renderTreeNodes = (nodes: ITreeNode[]) => {
      return (
        <ul className={styles.treeList}>
          {nodes.map(node => (
            <li key={node.key} className={styles.treeNode}>
              <div className={styles.nodeContent} onClick={() => this.handleNodeClick(node)}>
                {node.isFolder && (
                  <span className={styles.expanderIcon}>
                    {node.isExpanded ? '▼' : '►'}
                  </span>
                )}
                <i className={`ms-Icon ms-Icon--${node.icon}`} aria-hidden="true" style={{ marginRight: '5px' }}></i>
                <span>{escape(node.label)}</span>
              </div>
              {node.isFolder && node.isExpanded && node.children && node.children.length > 0 && (
                <div className={styles.childrenContainer}>
                  {renderTreeNodes(node.children)}
                </div>
              )}
               {node.isFolder && node.isExpanded && node.children && node.children.length === 0 && loading && (
                <div className={styles.loadingIndicator}>Carregando...</div>
              )}
            </li>
          ))}
        </ul>
      );
    };

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
            <p>
              {!selectedLibraryUrl
                ? "Por favor, abra as configurações da Web Part e selecione uma biblioteca de documentos."
                : (!metadataColumn1 && !metadataColumn2 && !metadataColumn3)
                  ? "Nenhuma coluna de metadados selecionada. Exibindo documentos da raiz da biblioteca (se houver)."
                  : "A biblioteca selecionada não contém documentos com os metadados especificados, ou ocorreu um erro."
              }
            </p>
          )}
          {!loading && !error && treeData.length > 0 && (
            <div>
              {renderTreeNodes(treeData)}
            </div>
          )}
        </div>
      </section>
    );
  }
}