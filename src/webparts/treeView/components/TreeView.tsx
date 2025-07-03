// src/webparts/treeView/components/TreeView.tsx

import * as React from 'react';
import styles from './TreeView.module.scss';
import { ITreeViewProps } from './ITreeViewProps'; // Sua interface de props

// Importações do PnP.js v1
import pnp from "sp-pnp-js";
// Removido o import "@pnp/sp/lists"; pois o método problemático não é mais usado aqui

// Para tipagem do contexto
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';


// Interfaces para a estrutura da árvore (nós)
interface ITreeNode {
  key: string; // Identificador único para o nó
  label: string; // Texto a ser exibido no nó
  icon?: string; // Nome do ícone Fluent UI (ex: 'Folder', 'Document')
  url?: string; // URL do documento (se for um arquivo)
  isFolder: boolean; // Indica se é uma pasta ou arquivo
  children?: ITreeNode[]; // Subnós (pastas/arquivos dentro)
  isExpanded?: boolean; // Estado de expansão do nó (para pastas)
  serverRelativeUrl?: string; // URL relativa ao servidor (usada para buscar subitens)
}

// Interface para o estado interno do componente TreeView
interface IComponentTreeViewState {
  treeData: ITreeNode[]; // Os dados da árvore
  loading: boolean; // Indicador de carregamento
  error: string; // Mensagem de erro, se houver
}

export default class TreeView extends React.Component<ITreeViewProps, IComponentTreeViewState> {
  constructor(props: ITreeViewProps) {
    super(props);
    this.state = {
      treeData: [],
      loading: true, // Começa carregando, mas pode mudar dependendo se uma URL é passada
      error: ""
    };
  }

  public async componentDidMount(): Promise<void> {
    // Carrega a biblioteca selecionada se houver uma URL, senão exibe uma mensagem
    if (this.props.selectedLibraryUrl) {
      await this.loadSelectedLibraryContents(this.props.selectedLibraryUrl, this.props.selectedLibraryTitle || "Biblioteca");
    } else {
      this.setState({
        loading: false,
        error: "Por favor, selecione uma biblioteca de documentos nas configurações da Web Part."
      });
    }
  }

  // Chamado quando as propriedades do componente são atualizadas (ex: usuário seleciona outra biblioteca)
  public async componentDidUpdate(prevProps: ITreeViewProps): Promise<void> {
    if (this.props.selectedLibraryUrl !== prevProps.selectedLibraryUrl) {
      if (this.props.selectedLibraryUrl) {
        await this.loadSelectedLibraryContents(this.props.selectedLibraryUrl, this.props.selectedLibraryTitle || "Biblioteca");
      } else {
        // Se a seleção foi limpa ou não há seleção, limpa a árvore e exibe mensagem
        this.setState({
          treeData: [],
          loading: false,
          error: "Por favor, selecione uma biblioteca de documentos nas configurações da Web Part."
        });
      }
    }
  }

  // NOVO MÉTODO: Carrega o conteúdo de UMA biblioteca selecionada
  // Agora recebe o título diretamente
  private async loadSelectedLibraryContents(libraryUrl: string, libraryTitle: string): Promise<void> {
    try {
      this.setState({ loading: true, error: "" });

      // REMOVIDA A LINHA QUE CAUSAVA O ERRO (getByRootFolderServerRelativeUrl)
      // Usaremos o título (libraryTitle) e a URL (libraryUrl) passados como propriedades diretamente.

      // Obtém os conteúdos da pasta raiz da biblioteca (pastas e arquivos de 1º nível)
      const rootContents = await this.getFolderContents(libraryUrl); // Reutiliza o método existente

      // Cria o nó raiz para a biblioteca selecionada
      const libraryRootNode: ITreeNode = {
        key: libraryUrl, // Usar a URL como chave, já que o ID da lista não está sendo buscado aqui
        label: libraryTitle,
        icon: "Library", // Ícone para bibliotecas
        isFolder: true,
        serverRelativeUrl: libraryUrl,
        children: rootContents, // Os filhos já são carregados no primeiro nível
        isExpanded: true // Expande a biblioteca automaticamente na carga inicial
      };

      this.setState({ treeData: [libraryRootNode], loading: false });

    } catch (error) {
      console.error("Erro ao carregar o conteúdo da biblioteca selecionada:", error);
      this.setState({ error: `Não foi possível carregar a biblioteca selecionada: ${escape(error.message)}`, loading: false, treeData: [] });
    }
  }

  // Método original para obter o conteúdo de uma pasta (subpastas e arquivos)
  // Reutilizado por loadSelectedLibraryContents e handleNodeClick
  private async getFolderContents(folderServerRelativeUrl: string): Promise<ITreeNode[]> {
    const nodes: ITreeNode[] = [];
    try {
      const folderContents = await pnp.sp.web.getFolderByServerRelativeUrl(folderServerRelativeUrl)
                                   .expand("Folders,Files")
                                   .get();

      for (const sub of folderContents.Folders) {
        nodes.push({
          key: sub.UniqueId,
          label: sub.Name,
          icon: "Folder",
          isFolder: true,
          serverRelativeUrl: sub.ServerRelativeUrl,
          children: [],
          isExpanded: false
        });
      }

      for (const file of folderContents.Files) {
        nodes.push({
          key: file.UniqueId,
          label: file.Name,
          icon: this.getFileIcon(file.Name),
          isFolder: false,
          url: file.ServerRelativeUrl,
          serverRelativeUrl: file.ServerRelativeUrl
        });
      }

    } catch (error) {
      console.error(`Erro ao carregar conteúdo da pasta ${folderServerRelativeUrl}:`, error);
      // Se houver erro ao carregar conteúdo de uma subpasta, podemos retornar vazia
    }
    return nodes;
  }

  private getFileIcon(fileName: string): string {
    const extension = fileName.split('.').pop()?.toLowerCase();
    switch (extension) {
      case 'docx': case 'doc': return 'WordDocument';
      case 'xlsx': case 'xls': return 'ExcelDocument';
      case 'pptx': case 'ppt': return 'PowerPointDocument';
      case 'pdf': return 'PDF';
      case 'jpg': case 'jpeg': case 'png': case 'gif': return 'Photo2';
      case 'zip': return 'ZipFolder';
      case 'txt': return 'TextDocument';
      default: return 'Document';
    }
  }

  private handleNodeClick = async (node: ITreeNode): Promise<void> => {
    if (node.isFolder) {
      this.setState(prevState => {
        const newTreeData = this.toggleNodeExpansion(prevState.treeData, node.key);
        return { treeData: newTreeData };
      }, async () => {
        const updatedNode = this.findNodeInTree(this.state.treeData, node.key);
        if (updatedNode && updatedNode.isExpanded && updatedNode.children && updatedNode.children.length === 0 && updatedNode.serverRelativeUrl) {
          this.setState({ loading: true });
          const children = await this.getFolderContents(updatedNode.serverRelativeUrl);
          this.setState(prevState => {
            const treeDataWithChildren = this.addChildrenToNode(prevState.treeData, updatedNode.key, children);
            return { treeData: treeDataWithChildren, loading: false };
          });
        }
      });
    } else if (node.url) {
      window.open(node.url, '_blank');
    }
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
          <h2>Visualizador de Pastas e Documentos</h2>
          <div>Web part property value: <strong>{escape(this.props.description)}</strong></div>
        </div>
        <div className={styles.treeContainer}>
          {loading && treeData.length === 0 && <p>Carregando...</p>}
          {error && <p style={{ color: 'red' }}>{error}</p>}
          {!loading && !error && treeData.length === 0 && (
            <p>
              {this.props.selectedLibraryUrl
                ? "A biblioteca selecionada não contém pastas ou documentos visíveis, ou ocorreu um erro."
                : "Por favor, abra as configurações da Web Part e selecione uma biblioteca de documentos."
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