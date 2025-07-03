
import * as React from 'react';
import styles from './TreeView.module.scss';
import { ITreeViewProps } from './ITreeViewProps'; // Sua interface de props

// Importações do PnP.js v1
import pnp from "sp-pnp-js";


import { WebPartContext } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';



interface ITreeNode {
  key: string; 
  label: string; 
  icon?: string; 
  url?: string; // URL do documento (se for um arquivo)
  isFolder: boolean; // Indica se é uma pasta ou arquivo
  children?: ITreeNode[]; // Subnós (pastas/arquivos dentro)
  isExpanded?: boolean; // Estado de expansão do nó (para pastas)
  serverRelativeUrl?: string; // URL relativa ao servidor (usada para buscar subitens)
}


interface IComponentTreeViewState {
  treeData: ITreeNode[]; 
  loading: boolean; 
  error: string; 
}

export default class TreeView extends React.Component<ITreeViewProps, IComponentTreeViewState> {
  constructor(props: ITreeViewProps) {
    super(props);
   
    this.state = {
      treeData: [],
      loading: true,
      error: ""
    };
  }


  public async componentDidMount(): Promise<void> {
    await this.loadDocumentLibraries(); // Inicia o carregamento das bibliotecas
  }

  // Método para carregar as bibliotecas de documentos do SharePoint
  private async loadDocumentLibraries(): Promise<void> {
    try {
      this.setState({ loading: true, error: "" }); 

     
      const allLists = await pnp.sp.web.lists
                                    .select("Title", "Id", "BaseTemplate", "Hidden", "RootFolder/ServerRelativeUrl")
                                    .expand("RootFolder")
                                    .get();

      const documentLibraries: ITreeNode[] = [];

      // Filtra apenas as bibliotecas de documentos (BaseTemplate === 101) e que não são ocultas
      for (const list of allLists) {
        if (list.BaseTemplate === 101 && !list.Hidden) {
          
          const rootFolderUrl = list.RootFolder.ServerRelativeUrl;
          documentLibraries.push({
            key: list.Id,
            label: list.Title,
            icon: "Library", // Ícone para bibliotecas
            isFolder: true,
            serverRelativeUrl: rootFolderUrl,
            children: [], 
            isExpanded: false
          });
        }
      }

      this.setState({ treeData: documentLibraries, loading: false }); // Atualiza o estado com as bibliotecas carregadas

    } catch (error) {
      console.error("Erro ao carregar bibliotecas de documentos:", error);
      this.setState({ error: "Não foi possível carregar as bibliotecas de documentos. Verifique as permissões ou se há bibliotecas no site.", loading: false });
    }
  }

 
  // Isso será chamado quando uma pasta for expandida
  private async getFolderContents(folderServerRelativeUrl: string): Promise<ITreeNode[]> {
    const nodes: ITreeNode[] = [];
    try {
      
      const folderContents = await pnp.sp.web.getFolderByServerRelativeUrl(folderServerRelativeUrl)
                                   .expand("Folders,Files") // Solicita que pastas e arquivos sejam incluídos na resposta
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
      
    }
    return nodes;
  }

  // Função auxiliar para determinar o ícone do arquivo com base na extensão
  private getFileIcon(fileName: string): string {
    const extension = fileName.split('.').pop()?.toLowerCase();
    switch (extension) {
      case 'docx':
      case 'doc': return 'WordDocument';
      case 'xlsx':
      case 'xls': return 'ExcelDocument';
      case 'pptx':
      case 'ppt': return 'PowerPointDocument';
      case 'pdf': return 'PDF';
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif': return 'Photo2';
      case 'zip': return 'ZipFolder';
      case 'txt': return 'TextDocument';
      default: return 'Document';
    }
  }

  // Lógica para lidar com o clique em qualquer nó da árvore
  private handleNodeClick = async (node: ITreeNode): Promise<void> => {
    if (node.isFolder) {
      // Se for uma pasta, alterna o estado de expansão
      this.setState(prevState => {
        const newTreeData = this.toggleNodeExpansion(prevState.treeData, node.key);
        return { treeData: newTreeData };
      }, async () => {
        // Callback após a atualização do estado: verifica se a pasta foi expandida
        const updatedNode = this.findNodeInTree(this.state.treeData, node.key);
        // Se a pasta foi expandida, ainda não tem filhos e tem um URL, carrega os filhos
        if (updatedNode && updatedNode.isExpanded && updatedNode.children && updatedNode.children.length === 0 && updatedNode.serverRelativeUrl) {
          this.setState({ loading: true }); // Define estado de carregamento para os filhos
          const children = await this.getFolderContents(updatedNode.serverRelativeUrl);
          this.setState(prevState => {
            const treeDataWithChildren = this.addChildrenToNode(prevState.treeData, updatedNode.key, children);
            return { treeData: treeDataWithChildren, loading: false }; // Atualiza com os filhos carregados
          });
        }
      });
    } else if (node.url) {
      // Se for um arquivo, abre a URL em uma nova aba do navegador
      window.open(node.url, '_blank');
    }
  };

  // Funções auxiliares para manipular a estrutura de dados da árvore (imutavelmente)
  // Alterna o estado 'isExpanded' de um nó específico
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

  // Adiciona um array de filhos a um nó pai específico
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

  // Encontra um nó na árvore recursivamente com base na sua chave
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

  // Método de renderização principal do componente
  public render(): React.ReactElement<ITreeViewProps> {
    const { loading, error, treeData } = this.state;
    
    // const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    
    const renderTreeNodes = (nodes: ITreeNode[]) => {
      return (
        <ul className={styles.treeList}> {/* Classe CSS para a lista da árvore */}
          {nodes.map(node => (
            <li key={node.key} className={styles.treeNode}> {/* Classe CSS para cada nó */}
              <div className={styles.nodeContent} onClick={() => this.handleNodeClick(node)}>
                {node.isFolder && (
                  <span className={styles.expanderIcon}>
                    {node.isExpanded ? '▼' : '►'} {/* Ícone de expansão/colapso para pastas */}
                  </span>
                )}
                {/* Ícone do Fluent UI baseado no tipo de nó */}
                <i className={`ms-Icon ms-Icon--${node.icon}`} aria-hidden="true" style={{ marginRight: '5px' }}></i>
                <span>{escape(node.label)}</span> {/* Exibe o label do nó, escapando HTML */}
              </div>
              {node.isFolder && node.isExpanded && node.children && node.children.length > 0 && (
                <div className={styles.childrenContainer}> {/* Contêiner para os filhos, com indentação */}
                  {renderTreeNodes(node.children)} {/* Chamada recursiva para renderizar os filhos */}
                </div>
              )}
               {node.isFolder && node.isExpanded && node.children && node.children.length === 0 && loading && (
                <div className={styles.loadingIndicator}>Carregando...</div> // Indicador de carregamento para pastas vazias
              )}
            </li>
          ))}
        </ul>
      );
    };

    return (
      <section className={styles.treeView}>
        <div className={styles.welcome}>
          <h2>Visualizador de Pastas e Documentos</h2>
          <div>Web part property value: <strong>{escape(this.props.description)}</strong></div>
        </div>
        <div className={styles.treeContainer}> {/* Contêiner principal da árvore */}
          {loading && treeData.length === 0 && <p>Carregando bibliotecas e pastas...</p>} {/* Mensagem de carregamento inicial */}
          {error && <p style={{ color: 'red' }}>{error}</p>} {/* Exibe mensagens de erro */}
          {!loading && !error && treeData.length === 0 && <p>Nenhuma biblioteca de documentos encontrada.</p>} {/* Mensagem se não houver dados */}
          {!loading && !error && treeData.length > 0 && (
            <div>
              {renderTreeNodes(treeData)} {/* Renderiza a árvore se houver dados */}
            </div>
          )}
        </div>
      </section>
    );
  }
}