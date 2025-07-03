// src/webparts/treeView/components/ITreeViewProps.ts
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITreeViewProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext; // <--- Certifique-se que esta linha está aqui!
  
}