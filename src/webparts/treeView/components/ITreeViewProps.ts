// src/webparts/treeView/components/ITreeViewProps.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ReactNode } from 'react';

export interface ITreeViewProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  selectedLibraryUrl?: string;
  selectedLibraryTitle?: string;
  metadataColumn1?: string; // Coluna de metadados para o Nível 1
  metadataColumn2?: string; // Coluna de metadados para o Nível 2
  metadataColumn3?: string; // Coluna de metadados para o Nível 3
  children?: ReactNode;
}