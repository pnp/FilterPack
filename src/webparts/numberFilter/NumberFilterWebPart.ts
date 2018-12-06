import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup,
} from '@microsoft/sp-webpart-base';

import * as strings from 'NumberFilterWebPartStrings';
import { INumberFilterProps, NumberFilter } from './components/NumberFilter';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { PropertyFieldSpinButton } from '@pnp/spfx-property-controls/lib/PropertyFieldSpinButton';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { sendValueAsPercentCallout, ratingIconCallout } from './components/PropertyCallouts';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';

export interface INumberFilterWebPartProps {
  defaultValue: number;
  displayAs: string;
  min: number;
  max: number;
  step: number;
  decimalPlaces: number;
  suffix: string;
  sendValueAsPercent: boolean;
  showSliderValue: boolean;
  sliderDirection: string;
  sliderHeight: number;
  sliderAlign: string;
  largeStars: boolean;
  //starIcon: string;
  labelShow: boolean;
  labelText: string;
  title: string;
  syncQS: boolean;
  QSkey: string;
}

export default class NumberFilterWebPart extends BaseClientSideWebPart<INumberFilterWebPartProps> implements IDynamicDataCallables {

  private _value: number;
  private _previousQSkey: string;

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);

    this._value = this.properties.defaultValue;
    this._previousQSkey = this.properties.QSkey;

    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      if(qsParams.has(this.properties.QSkey)) {
        this._value = Number(qsParams.get(this.properties.QSkey)) || 0;
      }
    }

    if(this._value < this.properties.min) {
      this._value = this.properties.min;
    }
    if(this._value > this.properties.max) {
      this._value = this.properties.max;
    }

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<INumberFilterProps > = React.createElement(
      NumberFilter,
      {
        value: this._value,
        displayAs: this.properties.displayAs,
        min: this.properties.min,
        max: this.properties.max,
        step: this.properties.step,
        decimalPlaces: this.properties.decimalPlaces,
        suffix: this.properties.suffix,
        showSliderValue: this.properties.showSliderValue,
        sliderVertical: this.properties.sliderDirection === "v",
        sliderHeight: this.properties.sliderHeight,
        sliderAlign: this.properties.sliderAlign,
        largeStars: this.properties.largeStars,
        //starIcon: this.properties.starIcon,
        labelShow: this.properties.labelShow,
        labelText: this.properties.labelText,
        onChange: (value: number) => {
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
      qsParams.set(this.properties.QSkey, this._value.toString());
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
        description: 'The value (number) of the filter',
      },
    ];
  }

  public getPropertyValue(propertyId: string): any {
    switch(propertyId) {
      case 'filterValue':
        return this.properties.sendValueAsPercent ? this._value/100 : this._value;
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
      PropertyPaneChoiceGroup('displayAs', {
        label: 'Display as',
        options: [
          {key: "slider", iconProps: {officeFabricIconFontName: 'Slider'}, text: 'Slider'},
          {key: "field", iconProps: {officeFabricIconFontName: 'NumberField'}, text: 'Field'},
          {key: "rating", iconProps: {officeFabricIconFontName: 'FavoriteStarFill'}, text: 'Rating'},
        ],
      })
    ];

    //Rating Specific Props
    if(this.properties.displayAs == "rating") {
      displayPaneFields.push(
        PropertyPaneToggle('largeStars', {
          label: "Rating icon size",
          onText: "Large",
          offText: "Small"
        })
      );
    }

    // Slider Specific Props
    if(this.properties.displayAs === "slider") {
      displayPaneFields.push(
        PropertyPaneDropdown('sliderDirection', {
          label: 'Slider style',
          options: [
            {key:'h', text: 'Horizontal'},
            {key:'v', text: 'Vertical'},
          ]
        })
      );

      if(this.properties.sliderDirection === "v") {
        displayPaneFields.push(
          PropertyFieldSpinButton('sliderHeight', {
            label: 'Slider height',
            initialValue: this.properties.sliderHeight,
            onPropertyChange: this.onPropertyPaneFieldChanged,
            properties: this.properties,
            disabled: false,
            suffix: ' px',
            min: 80,
            step: 10,
            key: 'sliderHeight'
          }),
          PropertyPaneDropdown('sliderAlign', {
            label: 'Slider alignment',
            options: [
              {key:'flex-start', text:'Left'},
              {key:'center', text:'Center'},
              {key:'flex-end', text:'Right'}
            ]
          })
        );
      }

      displayPaneFields.push(
        PropertyPaneToggle('showSliderValue', {
          label: 'Show value'
        })
      );
    }


    let valuePaneFields: Array<any> = [];

    if(this.properties.displayAs === "slider" || this.properties.displayAs === "field") {
      valuePaneFields.push(
        PropertyFieldSpinButton('min', {
          label: 'Minimum value',
          initialValue: this.properties.min,
          onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
          properties: this.properties,
          decimalPlaces: this.properties.decimalPlaces,
          key: 'min'
        })
      );
    }

    valuePaneFields.push(
      PropertyFieldSpinButton('max', {
        label: 'Maximum value',
        initialValue: this.properties.max,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        properties: this.properties,
        decimalPlaces: this.properties.displayAs === "rating" ? 0 : this.properties.decimalPlaces,
        key: 'max'
      })
    );

    if(this.properties.displayAs === "slider" || this.properties.displayAs === "field") {
      valuePaneFields.push(
        PropertyFieldSpinButton('step', {
          label: 'Step value',
          initialValue: this.properties.step,
          onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
          properties: this.properties,
          min: this.properties.decimalPlaces ? 1/(Math.pow(10,this.properties.decimalPlaces)) : 1,
          decimalPlaces: this.properties.decimalPlaces,
          key: 'step'
        }),
        PropertyFieldSpinButton('decimalPlaces', {
          label: 'Decimal places',
          initialValue: this.properties.decimalPlaces,
          onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
          properties: this.properties,
          min: 0,
          key: 'decimalPlaces'
        }),
      );
    }

    //Spin Button Specific Display Props
    if(this.properties.displayAs === "field") {
      valuePaneFields.push(
        PropertyPaneTextField('suffix', {
          label: 'Suffix'
        })
      );
    }


    valuePaneFields.push(
      PropertyFieldSpinButton('defaultValue', {
        label: 'Default value',
        initialValue: this.properties.defaultValue,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        properties: this.properties,
        decimalPlaces: this.properties.displayAs === "rating" ? 0 : this.properties.decimalPlaces,
        key: 'defaultValue'
      })
    );

    let advancedPaneFields: Array<any> = [];

    /*if(this.properties.displayAs === "rating") {
      advancedPaneFields.push(
        PropertyFieldTextWithCallout('starIcon', {
          calloutTrigger: CalloutTriggers.Hover,
          value: this.properties.starIcon,
          key: 'ratingIcon',
          label: 'Rating icon',
          calloutContent: ratingIconCallout(),
          calloutWidth: 200,
        })
      );
    }*/

    advancedPaneFields.push(
      PropertyFieldToggleWithCallout('sendValueAsPercent', {
        calloutTrigger: CalloutTriggers.Hover,
        key: 'sendValueAsPercent',
        label: 'Send value as percent',
        checked: this.properties.sendValueAsPercent,
        calloutContent: sendValueAsPercentCallout(),
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
              groupName: 'Value',
              groupFields: valuePaneFields,
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

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    if (propertyPath === 'min') {
      if(this.properties.defaultValue < this.properties.min) {
        this.properties.defaultValue = this.properties.min;
      }
      if(this._value < this.properties.min) {
        this._value = this.properties.min;
        this.context.dynamicDataSourceManager.notifyPropertyChanged('filterValue');
      }
      if(this.properties.min > this.properties.max) {
        this.properties.min = this.properties.max;
      }
    }

    if (propertyPath === 'max') {
      if(this.properties.defaultValue > this.properties.max) {
        this.properties.defaultValue = this.properties.max;
      }
      if(this._value > this.properties.max) {
        this._value = this.properties.max;
        this.context.dynamicDataSourceManager.notifyPropertyChanged('filterValue');
      }
      if(this.properties.max < this.properties.min) {
        this.properties.max = this.properties.min;
      }
    }

    if (propertyPath === 'syncQS' || propertyPath === 'QSkey') {
      this.syncQueryString();
    }

    if (propertyPath === 'sendValueAsPercent') {
      this.context.dynamicDataSourceManager.notifyPropertyChanged('filterValue');
    }
  }
}
