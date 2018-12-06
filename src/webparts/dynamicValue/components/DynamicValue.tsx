import { DisplayMode } from "@microsoft/sp-core-library";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import * as React from "react";

import styles from "./DynamicValue.module.scss";

export interface IDynamicValueProps {
  value?: any;
  displayMode: DisplayMode;
  updateTitle: (title: string) => void;
  title: string;
  isConfigured: boolean;
  onConfigure: () => void;

  labelShow: boolean;
  labelText: string;
  labelBold: boolean;
  labelPosition: string;

  displayType: string;
  displayBoolTrue: string;
  displayBoolFalse: string;
  displayTemplate: string;
  displayUndefinedValue: string;
}

export class DynamicValue extends React.Component<IDynamicValueProps, {}> {
  public render(): React.ReactElement<IDynamicValueProps> {
    const valueDisplay = this.formattedValue();
    return (
      <div className={ styles.dynamicValue }>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateTitle}/>
        {!this.props.isConfigured &&
          <Placeholder
            iconName='Edit'
            iconText='Setup the Dynamic Value'
            description='Choose a source and property to display'
            buttonLabel='Configure'
            onConfigure={this.props.onConfigure}/>
        }
        {this.props.isConfigured &&
          <div>
            {this.props.labelShow &&
              <span className={styles.label + (this.props.labelBold ? ' ' + styles.bold + ' ' : ' ') + (this.props.labelPosition=='beside' ? styles.beside : styles.above)}>
                {this.props.labelText}
              </span>
            }
            <span>
              {valueDisplay}
            </span>
          </div>
        }
      </div>
    );
  }

  private formattedValue(): string {
    let valueString: string;
    if(typeof this.props.value == "undefined") {
      valueString = this.props.displayUndefinedValue;
    } else {
      switch(this.props.displayType) {
        case "text":
          valueString = this.props.value;
          break;
        case "bool":
          valueString = this.props.value ? this.props.displayBoolTrue : this.props.displayBoolFalse;
          break;
      }
    }
    return this.props.displayTemplate.length ? this.props.displayTemplate.replace(/\[VALUE\]/g, valueString) : valueString;
  }
}
