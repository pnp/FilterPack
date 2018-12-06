import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './PeopleFilter.module.scss';
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPersonResult } from './IPersonResult';

export interface IPeopleFilterProps {
  selectedEmails: string[];
  context: WebPartContext;
  groupName?: string;
  showHiddenInUI: boolean;
  principalTypes: PrincipalType[];
  selectionLimit: number;
  useCurrentSite: boolean;
  webAbsoluteUrl: string;
  labelShow: boolean;
  labelText: string;
  onChange: (value: IPersonResult[]) => void;
  displayMode: DisplayMode;
  updateTitle: (title: string) => void;
  title: string;
}

export class PeopleFilter extends React.Component<IPeopleFilterProps, {}> {
  public render(): React.ReactElement<IPeopleFilterProps> {
    return (
      <div className={styles.peopleFilter}>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateTitle}/>
        <PeoplePicker
          context={this.props.context}
          defaultSelectedUsers={this.props.selectedEmails}
          titleText={this.props.labelShow ? this.props.labelText : undefined}
          peoplePickerWPclassName={!this.props.labelShow ? styles.noLabel : undefined}
          personSelectionLimit={this.props.selectionLimit}
          groupName={this.props.groupName}
          showHiddenInUI={this.props.showHiddenInUI}
          webAbsoluteUrl={this.props.useCurrentSite ? undefined : this.props.webAbsoluteUrl}
          selectedItems={(items: any[]) => {
            this.props.onChange(items.map(result => {
              return {
                id: parseInt(result.id),
                imageUrl: result.secondaryText.length ? result.imageUrl : "",
                email: result.secondaryText,
                title: result.text
              };
            }));
          }}
          principleTypes={this.props.principalTypes} />
      </div>
    );
  }
}
