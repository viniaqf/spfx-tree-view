// src/webparts/treeView/components/ITreeViewProps.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ReactNode } from 'react'; // Adicionado para 'children' em tipos mais recentes, embora possa não ser necessário explicitamente

export interface ITreeViewProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  selectedLibraryUrl?: string;
  selectedLibraryTitle?: string; 
  children?: ReactNode; // Adicionado para compatibilidade com tipos React Readonly<{ children?: ReactNode; }>
}