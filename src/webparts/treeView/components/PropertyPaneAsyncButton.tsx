import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
} from "@microsoft/sp-property-pane";
import {
  DefaultButton,
  Spinner,
  SpinnerSize,
  IButtonStyles,
  IIconProps,
  Label,
} from "office-ui-fabric-react";
import { getTranslations } from "../../../utils/getTranslations";


export interface IPropertyPaneAsyncButtonProps {
  label: string;
  isLoading: boolean;
  onClick: () => Promise<void>;
  disabled?: boolean;
}

interface IPropertyPaneAsyncButtonInternalProps
  extends IPropertyPaneAsyncButtonProps {
  onRender: (elem: HTMLElement) => void;
  onDispose?: (elem: HTMLElement) => void;
  key?: string;
}
const t = getTranslations();

class PropertyPaneAsyncButtonHost extends React.Component<IPropertyPaneAsyncButtonProps> {
  public render(): JSX.Element {
    const { label, isLoading, disabled, onClick } = this.props;

    const styles: IButtonStyles = {
      root: {
        width: "100%",
        height: "32px",
        marginTop: "10px",
      },
      label: {
        fontWeight: "bold",
      },
      icon: {
        margin: 0,
        marginLeft: 8,
      },
      flexContainer: {
        justifyContent: "center",
      },
    };

    return (
      <div style={{ padding: "0 18px", boxSizing: "border-box" }}>
        <DefaultButton
          styles={styles}
          disabled={disabled || isLoading}
          onClick={onClick}
        >
          {isLoading ? (
            <>
              {t.updating}
              <Spinner size={SpinnerSize.small} style={{ marginLeft: 8 }} />
            </>
          ) : (
            label
          )}
        </DefaultButton>
      </div>
    );
  }
}

export function PropertyPaneAsyncButton(
  targetProperty: string,
  properties: IPropertyPaneAsyncButtonProps
): IPropertyPaneField<IPropertyPaneAsyncButtonInternalProps> {
  return {
    type: PropertyPaneFieldType.Custom,
    targetProperty,
    properties: {
      key: `asyncButton_${targetProperty}`,
      ...properties,
      onRender: (elem: HTMLElement) => {
        const element: React.ReactElement<IPropertyPaneAsyncButtonProps> =
          React.createElement(PropertyPaneAsyncButtonHost, {
            ...properties,
          });
        ReactDom.render(element, elem);
      },
      onDispose: (elem: HTMLElement) => {
        try {
          ReactDom.unmountComponentAtNode(elem);
        } catch {
          // no-op
        }
      },
    },
  };
}
