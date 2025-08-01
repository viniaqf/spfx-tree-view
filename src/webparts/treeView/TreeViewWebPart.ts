// src/webparts/treeView/TreeViewWebPart.ts

import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "TreeViewWebPartStrings";
import TreeView from "./components/TreeView";
import { ITreeViewProps } from "./components/ITreeViewProps";
import pnp from "sp-pnp-js";

import { initializeIcons } from "office-ui-fabric-react";
import { getTranslations } from "../../utils/getTranslations";

export interface ITreeViewWebPartProps {
  description: string;
  selectedLibraryUrl?: string;
  selectedLibraryTitle?: string;
  metadataColumn1?: string;
  metadataColumn2?: string;
  metadataColumn3?: string;
  metadataColumnTypes?: {
    [internalName: string]: { type: string; lookupField?: string };
  }; // <-- Tipagem atualizada
}

const t = getTranslations();

export default class TreeViewWebPart extends BaseClientSideWebPart<ITreeViewWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  private _documentLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _metadataColumnOptions: IPropertyPaneDropdownOption[] = [];
  private _columnTypesMap: {
    [internalName: string]: { type: string; lookupField?: string };
  } = {}; // <-- Tipagem atualizada

  public render(): void {
    const element: React.ReactElement<ITreeViewProps> = React.createElement(
      TreeView,
      {
        //description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        selectedLibraryUrl: this.properties.selectedLibraryUrl,
        selectedLibraryTitle: this.properties.selectedLibraryTitle,
        metadataColumn1: this.properties.metadataColumn1,
        metadataColumn2: this.properties.metadataColumn2,
        metadataColumn3: this.properties.metadataColumn3,
        metadataColumnTypes: this._columnTypesMap,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // this._environmentMessage = this._getEnvironmentMessage();
    return super.onInit().then((_) => {
      pnp.setup({
        spfxContext: this.context,
      });
      initializeIcons();
    });
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    try {
      const libraries = await pnp.sp.web.lists
        .filter("BaseTemplate eq 101 and Hidden eq false")
        .select("Title", "Id", "RootFolder/ServerRelativeUrl")
        .expand("RootFolder")
        .get();

      this._documentLibraryOptions = libraries.map((lib) => ({
        key: lib.RootFolder.ServerRelativeUrl,
        text: lib.Title,
      }));
    } catch (error) {
      console.error(t.error_loading_library_options, error);
      this._documentLibraryOptions = [
        { key: "error", text: t.error_loading_libraries },
      ];
    }

    if (this.properties.selectedLibraryUrl) {
      try {
        const currentList = (
          await pnp.sp.web.lists
            .filter(
              `RootFolder/ServerRelativeUrl eq '${this.properties.selectedLibraryUrl}'`
            )
            .select("Id")
            .get()
        )[0];

        if (currentList && currentList.Id) {
          const rawListFields = await pnp.sp.web.lists
            .getById(currentList.Id)
            .fields.filter("Hidden eq false and ReadOnlyField eq false")
            .select("InternalName", "Title", "TypeAsString", "LookupField")
            .get();

          const allowedTypes = [
            "Text",
            "Note",
            "Number",
            "Integer",
            "DateTime",
            "Boolean",
            "Choice",
            "MultiChoice",
            "Lookup",
            "User",
            "ManagedMetadata",
          ];
          this._columnTypesMap = {}; // Inicializa o mapa de tipos
          this._metadataColumnOptions = rawListFields
            .filter((field) => allowedTypes.includes(field.TypeAsString))
            .map((field) => {
              let correctedInternalName = field.InternalName;
              if (
                (field.InternalName.endsWith("0") ||
                  field.InternalName.endsWith("_0")) &&
                (field.TypeAsString === "Lookup" ||
                  field.TypeAsString === "ManagedMetadata")
              ) {
                correctedInternalName = field.InternalName.substring(
                  0,
                  field.InternalName.length - 1
                );
                if (field.InternalName.endsWith("_0")) {
                  correctedInternalName = field.InternalName.substring(
                    0,
                    field.InternalName.length - 2
                  );
                }
              }

              this._columnTypesMap[correctedInternalName] = {
                type: field.TypeAsString,
                lookupField: field.LookupField || "Title",
              };

              return {
                key: correctedInternalName,
                text: field.Title,
              };
            });

          this._metadataColumnOptions.unshift({
            key: "",
            text: t.no_column,
          });
        } else {
          this._metadataColumnOptions = [
            {
              key: "",
              text: t.selected_library_not_found_or_missing_id,
            },
          ];
        }
      } catch (error) {
        console.error(t.error_loading_column_options, error);
        this._metadataColumnOptions = [
          { key: "error", text: t.error_loading_columns },
        ];
      }
    } else {
      this._metadataColumnOptions = [{ key: "", text: t.select_library_first }];
    }

    this.context.propertyPane.refresh();
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === "selectedLibraryUrl" && newValue !== oldValue) {
      const selectedOption = this._documentLibraryOptions.find(
        (option) => option.key === newValue
      );
      if (selectedOption) {
        this.properties.selectedLibraryTitle = selectedOption.text as string;
      } else {
        this.properties.selectedLibraryTitle = undefined;
      }
      // Limpa as colunas de metadados quando a biblioteca é alterada, evitando problemas de inconsistência
      this.properties.metadataColumn1 = "";
      this.properties.metadataColumn2 = "";
      this.properties.metadataColumn3 = "";

      this.onPropertyPaneConfigurationStart();
    }

    if (
      propertyPath === "selectedLibraryUrl" ||
      propertyPath === "metadataColumn1" ||
      propertyPath === "metadataColumn2" ||
      propertyPath === "metadataColumn3" ||
      propertyPath === "description"
    ) {
      this.render();
    }
  }

  // private _getEnvironmentMessage(): string {
  //   if (!!this.context.sdks.microsoftTeams) {
  //     return this.context.isServedFromLocalhost
  //       ? strings.AppLocalEnvironmentTeams
  //       : strings.AppTeamsTabEnvironment;
  //   }
  //   return this.context.isServedFromLocalhost
  //     ? strings.AppLocalEnvironmentSharePoint
  //     : strings.AppSharePointEnvironment;
  // }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    //Mantendo esse método para evitar possíveis erros de mudanças futuras.
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const disableColumnDropdowns =
      !this.properties.selectedLibraryUrl ||
      this._metadataColumnOptions.length === 0 ||
      (this._metadataColumnOptions.length === 1 &&
        this._metadataColumnOptions[0].key === "");

    const pnpV1SafeFields = this._metadataColumnOptions.filter(
      (opt) =>
        opt.key !== "error" &&
        opt.key !== "" &&
        opt.key !== t.selected_library_not_found_or_missing_id
    );

    return {
      pages: [
        {
          // header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              // groupName: strings.BasicGroupName,
              groupFields: [
                // PropertyPaneTextField('description', {
                //   label: strings.DescriptionFieldLabel
                // }),
                PropertyPaneDropdown("selectedLibraryUrl", {
                  label: t.select_library,
                  options: this._documentLibraryOptions,
                  selectedKey: this.properties.selectedLibraryUrl,
                  disabled: this._documentLibraryOptions.length === 0,
                }),
                PropertyPaneDropdown("metadataColumn1", {
                  label: t.metadata_column_level_1,
                  options: pnpV1SafeFields,
                  selectedKey: this.properties.metadataColumn1,
                  disabled: disableColumnDropdowns,
                }),
                PropertyPaneDropdown("metadataColumn2", {
                  label: t.metadata_column_level_2,
                  options: pnpV1SafeFields,
                  selectedKey: this.properties.metadataColumn2,
                  disabled:
                    disableColumnDropdowns || !this.properties.metadataColumn1,
                }),
                PropertyPaneDropdown("metadataColumn3", {
                  label: t.metadata_column_level_3,
                  options: pnpV1SafeFields,
                  selectedKey: this.properties.metadataColumn3,
                  disabled:
                    disableColumnDropdowns || !this.properties.metadataColumn2,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
