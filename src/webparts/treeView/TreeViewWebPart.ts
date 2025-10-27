import * as React from "react";
import TreeViewConfigService from "./services/TreeViewConfigService";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  PropertyPaneToggle,
  PropertyPaneButton,
  PropertyPaneButtonType,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "TreeViewWebPartStrings";
import TreeView from "./components/TreeView";
import { ITreeViewProps } from "./components/ITreeViewProps";
import pnp from "sp-pnp-js";

import { initializeIcons } from "office-ui-fabric-react";
import { getTranslations } from "../../utils/getTranslations";
import { PropertyPaneAsyncButton } from "./components/PropertyPaneAsyncButton";

export interface ITreeViewWebPartProps {
  description: string;
  selectedLibraryUrl?: string;
  selectedLibraryTitle?: string;
  metadataColumn1?: string;
  metadataColumn2?: string;
  metadataColumn3?: string;
  metadataColumnTypes?: {
    [internalName: string]: { type: string; lookupField?: string };
  };
  customLibraryTitlePT?: string;
  customLibraryTitleES?: string;
}

const t = getTranslations();

export default class TreeViewWebPart extends BaseClientSideWebPart<ITreeViewWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private _isRefreshing: boolean = false;

  private _documentLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _metadataColumnOptions: IPropertyPaneDropdownOption[] = [];
  private _columnTypesMap: {
    [internalName: string]: { type: string; lookupField?: string };
  } = {};

  public render(): void {
    const element: React.ReactElement<ITreeViewProps> = React.createElement(
      TreeView,
      {
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
        customLibraryTitlePT: this.properties.customLibraryTitlePT, //SNO365-89
        customLibraryTitleES: this.properties.customLibraryTitleES, //SNO365-89
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    pnp.setup({ spfxContext: this.context });
    initializeIcons();

    // URL absoluta para evitar colisão entre sites distintos com mesmo path
    const pageUrl = `${this.context.pageContext.web.absoluteUrl}${window.location.pathname}`;

    try {
      const cfg = await TreeViewConfigService.loadByPage(pageUrl);
      if (cfg) {
        // Só hidrata se as props ainda não foram definidas (evita sobrescrever edição em curso)
        const isEmpty =
          !this.properties.selectedLibraryUrl &&
          !this.properties.metadataColumn1 &&
          !this.properties.metadataColumn2 &&
          !this.properties.metadataColumn3;

        if (isEmpty) {
          this.properties.selectedLibraryUrl = cfg.Library || "";

          const h = cfg.Hierarchy
            ? (JSON.parse(cfg.Hierarchy) as string[])
            : [];
          this.properties.metadataColumn1 = h[0] || "";
          this.properties.metadataColumn2 = h[1] || "";
          this.properties.metadataColumn3 = h[2] || "";

          // Atualiza UI
          if (this.renderedOnce) this.render();
        }
      }
    } catch (e) {
      console.warn(
        "Falha ao hidratar propriedades a partir da lista TreeViewConfigs:",
        e
      );
    }
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

  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): Promise<void> {
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

      // Limpa as colunas quando a biblioteca muda
      this.properties.metadataColumn1 = "";
      this.properties.metadataColumn2 = "";
      this.properties.metadataColumn3 = "";

      // ⚠️ Aguarde recarregar opções antes de renderizar e salvar
      await this.onPropertyPaneConfigurationStart();

      // Renderiza depois que as opções foram atualizadas
      this.render();

      // Salva config mínima (sem o JSON publicado)
      try {
        const pageUrl = window.location.pathname;
        const hierarchy = JSON.stringify(
          [
            this.properties.metadataColumn1,
            this.properties.metadataColumn2,
            this.properties.metadataColumn3,
          ].filter(Boolean)
        );
        await TreeViewConfigService.upsert({
          PageURL: pageUrl,
          Library: this.properties.selectedLibraryUrl || "",
          Hierarchy: hierarchy,
        });
      } catch (e) {
        console.warn(
          "Não foi possível salvar config mínima na lista TreeViewConfigs:",
          e
        );
      }
      return;
    }

    if (
      propertyPath === "metadataColumn1" ||
      propertyPath === "metadataColumn2" ||
      propertyPath === "metadataColumn3" ||
      propertyPath === "description"
    ) {
      this.render();
      try {
        const pageUrl = window.location.pathname;
        const hierarchy = JSON.stringify(
          [
            this.properties.metadataColumn1,
            this.properties.metadataColumn2,
            this.properties.metadataColumn3,
          ].filter(Boolean)
        );
        await TreeViewConfigService.upsert({
          PageURL: pageUrl,
          Library: this.properties.selectedLibraryUrl || "",
          Hierarchy: hierarchy,
        });
      } catch (e) {
        console.warn(
          "Não foi possível salvar config mínima na lista TreeViewConfigs:",
          e
        );
      }
    }
  }

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

  private async handleRefreshClick(): Promise<void> {
    if (this._isRefreshing) {
      return;
    }

    this._isRefreshing = true;
    this.context.propertyPane.refresh();

    try {
      await this.onPropertyPaneConfigurationStart();

      this.render();

      console.log(
        "Web part forçada a recarregar. O cache será limpo e os dados buscados novamente."
      );
    } catch (error) {
      console.error("Falha ao recarregar os metadados:", error);
    } finally {
      this._isRefreshing = false;
      this.context.propertyPane.refresh();
    }
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
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown("selectedLibraryUrl", {
                  label: t.select_library,
                  options: this._documentLibraryOptions,
                  selectedKey: this.properties.selectedLibraryUrl,
                  disabled: this._documentLibraryOptions.length === 0,
                }),
                PropertyPaneTextField("customLibraryTitlePT", {
                  label: "Nome da biblioteca em PT",
                  description:
                    "Se for definido, este nome será exibido no lugar do título original da biblioteca em português.",
                  placeholder:
                    this.properties.selectedLibraryTitle ||
                    "Título original da biblioteca.",
                  maxLength: 100,
                  disabled: !this.properties.selectedLibraryUrl,
                }),
                PropertyPaneTextField("customLibraryTitleES", {
                  label: "Nome da biblioteca em ES",
                  description:
                    "Se for definido, este nome será exibido no lugar do título original da biblioteca em espanhol.",
                  placeholder:
                    this.properties.selectedLibraryTitle ||
                    "Título original da biblioteca.",
                  maxLength: 100,
                  disabled: !this.properties.selectedLibraryUrl,
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
                PropertyPaneAsyncButton("refreshButton", {
                  label: t.reloadMetadata,
                  isLoading: this._isRefreshing,
                  onClick: this.handleRefreshClick.bind(this),
                  disabled: !this.properties.selectedLibraryUrl,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
