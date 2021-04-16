import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart,  } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

import * as strings from 'ToggleFilterWebPartStrings';
import { IToggleFilterProps, ToggleFilter} from './components/ToggleFilter';
import { defaultStateCallout, sendValueAsStringCallout } from './components/PropertyCallouts';

export interface IToggleFilterWebPartProps {
  defaultState: boolean;
  onText: string;
  offText: string;
  labelShow: boolean;
  labelText: string;
  sendValueAsString: boolean;
  trueValue: string;
  falseValue: string;
  title: string;
  syncQS: boolean;
  QSkey: string;
}

export default class ToggleFilterWebPart extends BaseClientSideWebPart<IToggleFilterWebPartProps> implements IDynamicDataCallables {

  private _value: boolean;
  private _lastOnText: string;
  private _lastOffText: string;
  private _previousQSkey: string;

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);

    this._value = this.properties.defaultState;
    this._lastOnText = this.properties.onText;
    this._lastOffText = this.properties.offText;
    this._previousQSkey = this.properties.QSkey;

    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      if(qsParams.has(this.properties.QSkey)) {
        this._value = qsParams.get(this.properties.QSkey) == "1" ? true : false;
      }
    }

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IToggleFilterProps > = React.createElement(
      ToggleFilter,
      {
        value: this._value,
        onText: this.properties.onText,
        offText: this.properties.offText,
        labelShow: this.properties.labelShow,
        labelText: this.properties.labelText,
        onChange: (value: boolean) => {
          this._value = value;
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterValue');
          this.syncQueryString();
          this.render();
        },
        displayMode: this.displayMode,
        updateTitle: (title: string) => {
          this.properties.title = title;
        },
        title: this.properties.title,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private syncQueryString(): void {
    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      qsParams.set(this.properties.QSkey, this._value ? "1" : "0");
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
        id: 'filterValue',
        title: 'Filter Value',
        description: 'The value (true/false) of the toggle',
      },
    ];
  }

  public getPropertyValue(propertyId: string): any {
    switch(propertyId) {
      case 'filterValue':
        return this.properties.sendValueAsString ? (this._value ? this.properties.trueValue : this.properties.falseValue) : this._value;
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
      PropertyPaneTextField('onText', {
        label: 'On Text'
      }),
      PropertyPaneTextField('offText', {
        label: 'Off Text'
      }),
      PropertyFieldToggleWithCallout('defaultState', {
        calloutTrigger: CalloutTriggers.Hover,
        key: 'defaultState',
        label: 'Default State',
        onText: this.properties.onText,
        offText: this.properties.offText,
        checked: this.properties.defaultState,
        calloutContent: defaultStateCallout(),
        calloutWidth: 200,
      }),
      PropertyFieldToggleWithCallout('sendValueAsString', {
        calloutTrigger: CalloutTriggers.Hover,
        key: 'sendValueAsString',
        label: 'Send value as text',
        checked: this.properties.sendValueAsString,
        calloutContent: sendValueAsStringCallout(),
        calloutWidth: 200,
      })
    ];

    if(this.properties.sendValueAsString) {
      displayPaneFields.push(...[
        PropertyPaneTextField('trueValue', {
          label: 'Text value when true',
        }),
        PropertyPaneTextField('falseValue', {
          label: 'Text value when false',
        })
      ]);
    }

    let qsPaneFields: Array<any> = [
      PropertyPaneToggle('syncQS', {
        label: 'Sync with Query String'
      })
    ];

    if(this.properties.syncQS) {
      qsPaneFields.push(
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
              groupFields: qsPaneFields,
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    let shouldNotify: boolean = false;
    //If true/false values are equal to on/off text then keep them in sync
    // Keeps them automatic but keeps any customizations
    if (propertyPath === 'onText') {
      if (typeof this.properties.trueValue == "undefined" || this.properties.trueValue == this._lastOnText) {
        this.properties.trueValue = this.properties.onText;
        shouldNotify = true;
      }
      this._lastOnText = this.properties.onText;
    }
    if (propertyPath === 'offText') {
      if (typeof this.properties.falseValue == "undefined" || this.properties.falseValue == this._lastOffText) {
        this.properties.falseValue = this.properties.offText;
        shouldNotify = true;
      }
      this._lastOffText = this.properties.offText;
    }
    if (propertyPath === 'sendValueAsString' || propertyPath === 'trueValue' || propertyPath === 'falseValue') {
      shouldNotify = true;
    }

    if (propertyPath === 'syncQS' || propertyPath === 'QSkey') {
      this.syncQueryString();
    }

    if(shouldNotify) {
      this.context.dynamicDataSourceManager.notifyPropertyChanged('filterValue');
    }
  }
}
