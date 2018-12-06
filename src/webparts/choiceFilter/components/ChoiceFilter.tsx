import * as React from 'react';
import styles from './ChoiceFilter.module.scss';
import { DisplayMode } from "@microsoft/sp-core-library";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface IChoiceFilterProps {
  displayMode: DisplayMode;
  updateTitle: (title: string) => void;
  title: string;
  isConfigured: boolean;
  onConfigure: () => void;

  labelShow: boolean;
  labelText: string;

  isLoading: boolean;
  options: IDropdownOption[];
  selectedKey?: any;
  allowNone: boolean;
  onChange: (option:IDropdownOption) => void;
  onBlank: () => void;
}


export class ChoiceFilter extends React.Component<IChoiceFilterProps, {}> {
  public render(): React.ReactElement<IChoiceFilterProps> {
    let key: any = this.props.selectedKey;
    if(typeof this.props.selectedKey == "undefined" && !this.props.allowNone) {
      if (this.props.options.length) {
        key = this.props.options[0].key;
        this.props.onChange(this.props.options[0]);
      }
    }

    return (
      <div className={ styles.choiceFilter }>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateTitle}/>
        {!this.props.isConfigured &&
          <Placeholder
            iconName='Edit'
            iconText='Setup the Choice Filter'
            description='Configure required options'
            buttonLabel='Configure'
            onConfigure={this.props.onConfigure}/>
        }
        {this.props.isConfigured && this.props.isLoading &&
          <Spinner
            size={SpinnerSize.small}/>
        }
        {this.props.isConfigured && !this.props.isLoading &&
          <div className={styles.propAndButtonBox}>
            <div className={styles.mainBox}>
              <Dropdown
                key={key}
                label={this.props.labelShow ? this.props.labelText : undefined}
                options={this.props.options}
                selectedKey={key}
                onChanged={this.props.onChange}/>
            </div>
            <div className={styles.buttonBox}>
              {this.props.allowNone &&
                <IconButton
                  iconProps={{iconName:'Clear'}}
                  disabled={typeof this.props.selectedKey == "undefined"}
                  onClick={this.props.onBlank}/>
              }
            </div>
          </div>
        }
      </div>
    );
  }
}
