import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'TextFilterWebPartStrings';
import {TextFilter, ITextFilterProps} from './components/TextFilter';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

export interface ITextFilterWebPartProps {
  defaultValue: string;
  placeholder: string;
  labelShow: boolean;
  labelText: string;
  title: string;
  syncQS: boolean;
  QSkey: string;
}

export default class TextFilterWebPart extends BaseClientSideWebPart<ITextFilterWebPartProps> implements IDynamicDataCallables {

  private _value: string;
  private _previousQSkey: string;

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);

    this._value = this.properties.defaultValue;
    this._previousQSkey = this.properties.QSkey;

    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      if(qsParams.has(this.properties.QSkey)) {
        this._value = qsParams.get(this.properties.QSkey);
      }
    }

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<ITextFilterProps > = React.createElement(
      TextFilter,
      {
        value: this._value,
        placeholder: this.properties.placeholder,
        labelShow: this.properties.labelShow,
        labelText: this.properties.labelText,
        onChange: (value: string) => {
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
      qsParams.set(this.properties.QSkey, this._value);
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
        description: 'The value of the text field',
      },
    ];
  }

  public getPropertyValue(propertyId: string): any {
    switch(propertyId) {
      case 'filterValue':
        return this._value;
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
      PropertyPaneTextField('defaultValue', {
        label: 'Default value'
      }),
      PropertyPaneTextField('placeholder', {
        label: 'Placeholder text'
      })
    ];

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
    if (propertyPath === 'syncQS' || propertyPath === 'QSkey') {
      this.syncQueryString();
    }
  }
}
