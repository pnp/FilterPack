import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
} from '@microsoft/sp-webpart-base';

import * as strings from 'PeopleFilterWebPartStrings';
import { PeopleFilter, IPeopleFilterProps} from './components/PeopleFilter';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import {PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPersonResult } from './components/IPersonResult';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { defaultValueCallout, showHiddenInUICallout, groupNameCallout, useCurrentSiteCallout, webAbsoluteUrlCallout } from './components/PropertyCallouts';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldSpinButton } from '@pnp/spfx-property-controls/lib/PropertyFieldSpinButton';

export interface IPeopleFilterWebPartProps {
  defaultValue: string;
  defaultCurrentUser: boolean;
  groupName: string;
  showHiddenInUI: boolean;
  principleTypes: Array<PrincipalType>;
  selectionLimit: number;
  useCurrentSite: boolean;
  webAbsoluteUrl: string;
  labelShow: boolean;
  labelText: string;
  title: string;
  syncQS: boolean;
  QSkey: string;
}

export default class PeopleFilterWebPart extends BaseClientSideWebPart<IPeopleFilterWebPartProps> implements  IDynamicDataCallables {

  private _value: IPersonResult[];
  private _previousQSkey: string;

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);

    if(this.properties.defaultCurrentUser) {
      this._value = [{email:this.context.pageContext.user.email}];
    } else {
      this._value = this.properties.defaultValue ? this.properties.defaultValue.split(';').map(value => {
          return {
            email: value
          };
        }) : [];
    }
    this._previousQSkey = this.properties.QSkey;

    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      if(qsParams.has(this.properties.QSkey)) {
        this._value = qsParams.get(this.properties.QSkey).split(';').map(value => {
          return {
            email: value
          };
        });
      }
    }

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IPeopleFilterProps > = React.createElement(
      PeopleFilter,
      {
        selectedEmails: this._value.map(value => {
          return value.email;
        }),
        context: this.context,
        groupName: this.properties.groupName,
        showHiddenInUI: this.properties.showHiddenInUI,
        principalTypes: this.properties.principleTypes,
        selectionLimit: this.properties.selectionLimit,
        useCurrentSite: this.properties.useCurrentSite,
        webAbsoluteUrl: this.properties.webAbsoluteUrl,
        labelShow: this.properties.labelShow,
        labelText: this.properties.labelText,
        displayMode: this.displayMode,
        updateTitle: (title: string) => {
          this.properties.title = title;
        },
        onChange: (results: IPersonResult[]) => {
          this._value = results;
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterId');
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterImageUrl');
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterEmail');
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterTitle');
          this.syncQueryString();
          //this.render();
        },
        title: this.properties.title,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private syncQueryString(): void {
    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      qsParams.set(this.properties.QSkey, this._value.map(value => {
        return value.email;
      }).join(';'));
      if (this.properties.QSkey !== this._previousQSkey && this._previousQSkey) {
        if (qsParams.has(this._previousQSkey)) {
          qsParams.delete(this._previousQSkey);
        }
        this._previousQSkey = this.properties.QSkey;
      }
      window.history.replaceState({},'',`${location.pathname}?${qsParams}`);
    }
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'filterId',
        title: 'User ID',
        description: 'The user ID(s) from the people filter',
      },
      {
        id: 'filterImageUrl',
        title: 'Image URL',
        description: 'The profile Image URL(s) from the people filter',
      },
      {
        id: 'filterEmail',
        title: 'Email',
        description: 'The email address(es) from the people filter',
      },
      {
        id: 'filterTitle',
        title: 'Title',
        description: 'The name(s) from the people filter',
      },
    ];
  }

  public getPropertyValue(propertyId: string): any {
    switch(propertyId) {
      case 'filterId':
        return this._value.map(value => {
            return value.id;
          }).join(';');
      case 'filterImageUrl':
        return this._value.map(value => {
            return value.imageUrl;
          }).join(';');
      case 'filterEmail':
        return this._value.map(value => {
            return value.email;
          }).join(';');
      case 'filterTitle':
        return this._value.map(value => {
            return value.title;
          }).join(';');
    }
    throw new Error('Bad property id');
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let labelPaneFields: Array<any> = [
      PropertyPaneToggle('labelShow', {
        label: 'Show Label',
      })
    ];

    if(this.properties.labelShow){
      labelPaneFields.push(
        PropertyPaneTextField('labelText', {
          label: 'Label Text',
        })
      );
    }

    let displayPaneFields: Array<any> = [
      PropertyPaneToggle('defaultCurrentUser', {
        label: 'Use the "current user" as the default'
      })
    ];

    if(!this.properties.defaultCurrentUser) {
      displayPaneFields.push(
        PropertyFieldTextWithCallout('defaultValue', {
          calloutTrigger: CalloutTriggers.Hover,
          value: this.properties.defaultValue,
          key: 'defaultValue',
          label: 'Default value(s)',
          calloutContent: defaultValueCallout(),
          calloutWidth: 200,
        })
      );
    }

    displayPaneFields.push(
      PropertyFieldMultiSelect('principleTypes', {
        key: 'principleTypes',
        label: 'Principal types to include',
        options: [
          {key:PrincipalType.User, text: 'Users'},
          {key:PrincipalType.SharePointGroup, text: 'SharePoint Groups'},
          {key:PrincipalType.SecurityGroup, text: 'Security Groups'},
          {key:PrincipalType.DistributionList, text: 'Distribution Lists'},
        ],
        selectedKeys: this.properties.principleTypes,
      }),
      PropertyFieldSpinButton('selectionLimit', {
        label: 'Selection limit',
        initialValue: this.properties.selectionLimit,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
          properties: this.properties,
          min: 1,
          key: 'selectionLimit'
      })
    );

    let advancedPaneFields: Array<any> = [
      PropertyFieldToggleWithCallout('useCurrentSite', {
        calloutTrigger: CalloutTriggers.Hover,
        checked: this.properties.useCurrentSite,
        key: 'useCurrentSite',
        label: 'Pull users from current site',
        calloutContent: useCurrentSiteCallout(),
        calloutWidth: 200,
      })
    ];

    if(!this.properties.useCurrentSite) {
      advancedPaneFields.push(
        PropertyFieldTextWithCallout('webAbsoluteUrl', {
          calloutTrigger: CalloutTriggers.Hover,
          value: this.properties.webAbsoluteUrl,
          key: 'webAbsoluteUrl',
          label: 'Web absolute URL',
          calloutContent: webAbsoluteUrlCallout(),
          calloutWidth: 260,
        })
      );
    }

    advancedPaneFields.push(
      PropertyFieldTextWithCallout('groupName', {
        calloutTrigger: CalloutTriggers.Hover,
        value: this.properties.groupName,
        key: 'groupName',
        label: 'Filter to group',
        calloutContent: groupNameCallout(),
        calloutWidth: 200,
      }),
      PropertyFieldToggleWithCallout('showHiddenInUI', {
        calloutTrigger: CalloutTriggers.Hover,
        checked: this.properties.showHiddenInUI,
        key: 'showHiddenInUI',
        label: 'Include hidden users/groups',
        calloutContent: showHiddenInUICallout(),
        calloutWidth: 200,
      }),
      PropertyPaneToggle('syncQS', {
        label: 'Sync with Query String'
      })
    );

    if(this.properties.syncQS) {
      advancedPaneFields.push(
        PropertyPaneTextField('QSkey', {
          label: 'Querystring key'
        })
      );
    }

    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Label',
              groupFields: labelPaneFields,
            },
            {
              groupName: 'Display',
              groupFields: displayPaneFields,
            },
            {
              groupName: 'Advanced',
              groupFields: advancedPaneFields,
            }
          ]
        }
      ]
    };
  }
}
