import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { TextField } from 'office-ui-fabric-react/lib/TextField';

export interface ITextFilterProps {
  value: string;
  placeholder: string;
  labelShow: boolean;
  labelText: string;
  onChange: (value: string) => void;
  displayMode: DisplayMode;
  updateTitle: (title: string) => void;
  title: string;
}

export class TextFilter extends React.Component<ITextFilterProps, {}> {
  public render(): React.ReactElement<ITextFilterProps> {
    return (
      <div>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateTitle}/>
        <TextField
          value={this.props.value}
          onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => this.props.onChange(newValue)}
          label={this.props.labelShow ? this.props.labelText : undefined}
          placeholder={this.props.placeholder}/>
      </div>
    );
  }
}
