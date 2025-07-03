// src/webparts/treeView/TreeViewWebPart.ts

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TreeViewWebPartStrings';
import TreeView from './components/TreeView';
import { ITreeViewProps } from './components/ITreeViewProps';
import pnp from "sp-pnp-js";

import { initializeIcons } from '@fluentui/react';

export interface ITreeViewWebPartProps {
  description: string;
  selectedLibraryUrl?: string;
}

export default class TreeViewWebPart extends BaseClientSideWebPart<ITreeViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _documentLibraryOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<ITreeViewProps> = React.createElement(
      TreeView,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        selectedLibraryUrl: this.properties.selectedLibraryUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
      initializeIcons();
    });
  }

  // Correção aqui: Remove as referências a 'loadingElement'
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    // this.context.propertyPane.loadingElement.innerHTML = 'Carregando bibliotecas de documentos...'; // <-- REMOVIDO

    try {
      const libraries = await pnp.sp.web.lists
                                    .filter("BaseTemplate eq 101 and Hidden eq false")
                                    .select("Title", "Id", "RootFolder/ServerRelativeUrl")
                                    .expand("RootFolder")
                                    .get();

      this._documentLibraryOptions = libraries.map(lib => ({
        key: lib.RootFolder.ServerRelativeUrl,
        text: lib.Title
      }));

    } catch (error) {
      console.error("Erro ao carregar opções de biblioteca:", error);
      this._documentLibraryOptions = [{ key: "error", text: "Erro ao carregar bibliotecas" }];
    } finally {
      // this.context.propertyPane.loadingElement.innerHTML = ''; // <-- REMOVIDO
      this.context.propertyPane.refresh(); // O PropertyPane já tem um spinner de carregamento automático
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'selectedLibraryUrl' && newValue) {
        this.render();
    }
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }
    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('selectedLibraryUrl', {
                  label: 'Selecionar Biblioteca',
                  options: this._documentLibraryOptions,
                  selectedKey: this.properties.selectedLibraryUrl,
                  // placeholder: 'Selecione uma biblioteca...', // <-- REMOVIDO
                  disabled: this._documentLibraryOptions.length === 0
                })
              ]
            }
          ]
        }
      ]
    };
  }
}