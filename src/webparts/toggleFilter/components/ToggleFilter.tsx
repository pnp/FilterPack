import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import * as React from "react";

export interface IToggleFilterProps {
  value: boolean;
  onText: string;
  offText: string;
  labelShow: boolean;
  labelText: string;
  onChange: (value: boolean) => void;
  displayMode: DisplayMode;
  updateTitle: (title: string) => void;
  title: string;
}

export class ToggleFilter extends React.Component<IToggleFilterProps, {}> {
  public render(): React.ReactElement<IToggleFilterProps> {
    return (
      <div>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateTitle}/>
        <Toggle
          checked={this.props.value}
          onChanged={this.props.onChange}
          label={this.props.labelShow ? this.props.labelText : undefined}
          onText={this.props.onText}
          offText={this.props.offText}/>
      </div>
    );
  }
}
