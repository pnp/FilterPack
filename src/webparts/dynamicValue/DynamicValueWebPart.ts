import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Log } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, IPropertyPaneDropdownOption, PropertyPaneDropdown, PropertyPaneToggle, PropertyPaneDynamicField } from "@microsoft/sp-property-pane";

import { IDynamicDataSource } from '@microsoft/sp-dynamic-data';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
import { PropertyFieldSpinButton } from '@pnp/spfx-property-controls/lib/PropertyFieldSpinButton';

import * as strings from 'DynamicValueWebPartStrings';
import { DynamicValue, IDynamicValueProps } from './components/DynamicValue';
import { displayTemplateCallout } from './components/PropertyCallouts';

export interface IDynamicValueWebPartProps {
  title: string;
  propertyId: string;
  sourceId: string;
  propertyTitle: string;
  labelShow: boolean;
  labelText: string;
  labelBold: boolean;
  labelPosition: string;
  displayType: string;
  displayBoolTrue: string;
  displayBoolFalse: string;
  displayObjectProperty: string;
  displayObjectPropertyManual: string;
  displaySubPropertyType: string;
  displayTemplate: string;
  displayUndefined: string;
  displayUndefinedCustom: string;
  displayArrayStyle: string;
  displayArrayIndex: number;
  displayArrayValueType: string;
}

export default class DynamicValueWebPart extends BaseClientSideWebPart<IDynamicValueWebPartProps> {

  private _lastSourceId: string;
  private _lastPropertyId: string;
  private _lastPropertyTitle: string;

  private _sourceAttemptCount: number;
  private _registrationComplete: boolean;

  protected onInit(): Promise<void> {
    this.render = this.render.bind(this);
    this._sourceAttemptCount = 0;
    this._registrationComplete = false;

    return Promise.resolve();
  }

  public render(): void {
    let value: any = undefined;
    const needsConfiguration: boolean = !(this.properties.sourceId && this.properties.propertyId);

    if(!needsConfiguration) {
      //Get the value from our source
      const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(this.properties.sourceId);
      if(!source) {
        this._sourceAttemptCount += 1;
        if(this._sourceAttemptCount <=3) {
          //console.log('retrying...');
          window.setTimeout(this.render,100);
        } else {
          this.context.statusRenderer.renderError(this.domElement, `Unable to connect to the data source (${this.properties.sourceId})`);
        }
        return;
      } else {
        this._sourceAttemptCount = 0;
      }
      try {
        value = source.getPropertyValue(this.properties.propertyId);
        if(typeof value !== "undefined" && this.properties.displayType === 'object') {
          value = value[this.properties.displayObjectProperty === 'manuallyspecified' ? this.properties.displayObjectPropertyManual : this.properties.displayObjectProperty];
        }
        if(typeof value !== "undefined" && this.properties.displayType === 'array') {
          switch(this.properties.displayArrayStyle) {
            case "index":
              if((value as any).length > this.properties.displayArrayIndex) {
                value = value[this.properties.displayArrayIndex];
                if(typeof value !== "undefined" && this.properties.displayArrayValueType === 'object') {
                  value = value[this.properties.displayObjectProperty === 'manuallyspecified' ? this.properties.displayObjectPropertyManual : this.properties.displayObjectProperty];
                }
              } else {
                value = undefined;
              }
          }
        }
      }
      catch(e) {
        const errorMessage = `An error has occurred while retrieving the property value (${this.properties.propertyId}). Details: ${e}`;
        this.context.statusRenderer.renderError(this.domElement, errorMessage);
        this.logError(errorMessage, e);
        return;
      }
    }

    if (this._registrationComplete === false && !needsConfiguration) {
      try {
        //If the subscription props exist, but we haven't registered yet then we need to register
        // otherwise, the registration will happen with the propertypane as the properties are configured
        this.context.dynamicDataProvider.registerPropertyChanged(this.properties.sourceId, this.properties.propertyId, this.render);
        this._lastSourceId = this.properties.sourceId;
        this._lastPropertyId = this.properties.propertyId;
        this._registrationComplete = true;
      }
      catch(e) {
        const errorMessage = `An error has occurred while connecting to the data source. Details: ${e}`;
        this.context.statusRenderer.renderError(this.domElement, errorMessage);
        this.logError(errorMessage, e);
        return;
      }
    }

    const element: React.ReactElement<IDynamicValueProps > = React.createElement(
      DynamicValue,
      {
        value: value,
        displayMode: this.displayMode,
        updateTitle: (title: string) => {
          this.properties.title = title;
        },
        title: this.properties.title,
        isConfigured: !needsConfiguration,
        onConfigure: () => {
          this.context.propertyPane.open();
        },
        labelShow: this.properties.labelShow,
        labelText: this.properties.labelText,
        labelBold: this.properties.labelBold,
        labelPosition: this.properties.labelPosition,
        displayType: this.properties.displayType === 'object' || (this.properties.displayType === 'array' && this.properties.displayArrayValueType === 'object') ? this.properties.displaySubPropertyType : this.properties.displayType,
        displayBoolTrue: this.properties.displayBoolTrue,
        displayBoolFalse: this.properties.displayBoolFalse,
        displayTemplate: this.properties.displayTemplate,
        displayUndefinedValue: this.properties.displayUndefined === 'blank' ? '' : (this.properties.displayUndefined === 'undefined' ? 'undefined' : this.properties.displayUndefinedCustom),
        //displayArrayStyle: this.properties.displayArrayStyle,
        //displayArrayIndex: this.properties.displayArrayIndex,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private logError(message:string, e:any) {
    Log.error(message,e,this.context.serviceScope);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const sourceOptions: IPropertyPaneDropdownOption[] = this.context.dynamicDataProvider.getAvailableSources()
      .map(source => {
        return {
          key: source.id,
          text: source.metadata.title,
        };
      });

    const selectedSource: string = this.properties.sourceId;
    let propertyOptions: IPropertyPaneDropdownOption[] = [];
    if(selectedSource) {
      const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(selectedSource);
      if(source) {
        propertyOptions = source.getPropertyDefinitions().map(prop => {
          return {
            key: prop.id,
            text: prop.title,
          };
        });
      }
    }

    let labelPaneFields: Array<any> = [
      PropertyPaneToggle('labelShow', {
        label: 'Show Label',
      })
    ];

    if(this.properties.labelShow){
      labelPaneFields.push(...[
        PropertyPaneTextField('labelText', {
          label: 'Label Text',
        }),
        PropertyPaneToggle('labelBold', {
          label: 'Bold',
        }),
        PropertyPaneDropdown('labelPosition', {
          label: 'Position',
          options: [
            {key:'above', text:'Above'},
            {key:'beside', text:'Beside'},
          ]
        })
      ]);
    }

    //Display Type
    let displayPaneFields: Array<any> = [
      PropertyPaneDropdown('displayType', {
        label: 'Display value as',
        options: [
          {key:'text', text:'Text'},
          {key:'bool', text:'Yes/No'},
          {key:'array', text:'Array'},
          {key:'object', text:'Object'},
        ]
      })
    ];

    if(this.properties.displayType === 'array') {
      displayPaneFields.push(
        PropertyPaneDropdown('displayArrayStyle', {
          label: 'Value By',
          options: [
            {key:'index', text:'Index'},
          ],
          selectedKey: this.properties.displayObjectProperty,
        })
      );

      if(this.properties.displayArrayStyle === 'index') {
        displayPaneFields.push(
          PropertyFieldSpinButton('displayArrayIndex', {
            label: 'Index',
            min: 0,
            key:'displayArrayIndex',
            properties: this.properties,
            onPropertyChange: this.onPropertyPaneFieldChanged,
          }),
          PropertyPaneDropdown('displayArrayValueType', {
            label: 'Display value as',
            options: [
              {key:'text', text:'Text'},
              {key:'bool', text:'Yes/No'},
              {key:'object', text:'Object'},
            ]
          })
        );
      }
    }

    if(this.properties.displayType === 'object' || this.properties.displayArrayValueType === 'object') {
      let keyOptions: IPropertyPaneDropdownOption[] = [];
      const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(this.properties.sourceId);
      if(source) {
        try {
          let sample: any = source.getPropertyValue(this.properties.propertyId);
          if(this.properties.displayType === 'array') {
            switch(this.properties.displayArrayStyle) {
              case "index":
                if(sample.length > this.properties.displayArrayIndex) {
                  sample = sample[this.properties.displayArrayIndex];
                }
            }
          }
          keyOptions = Object.keys(sample).map(key => {
            return {key:key, text:key};
          });
        }
        catch(e) {
        }
      }
      keyOptions.push({key:'manuallyspecified', text: 'Other'});

      //If the current displayObjectProperty isn't in the list, then pick the first one
      if(typeof this.properties.displayObjectProperty == "undefined" || !keyOptions.filter((opt:IPropertyPaneDropdownOption) => {
        return opt.key.toString() === this.properties.displayObjectProperty;
      }).length) {
        this.properties.displayObjectProperty = keyOptions[0].key.toString();
        this.render();
      }

      displayPaneFields.push(
        PropertyPaneDropdown('displayObjectProperty', {
          label: 'Sub property',
          options: keyOptions,
          selectedKey: this.properties.displayObjectProperty,
        })
      );

      if(this.properties.displayObjectProperty === 'manuallyspecified') {
        displayPaneFields.push(
          PropertyPaneTextField('displayObjectPropertyManual', {
            label: 'Sub Property Name'
          })
        );
      }

      displayPaneFields.push(
        PropertyPaneDropdown('displaySubPropertyType', {
          label: 'Display sub property as',
          options: [
            {key:'text', text:'Text'},
            {key:'bool', text:'Yes/No'},
          ],
        })
      );
    }

    //Optional configuration props depending on display type
    switch(this.properties.displayType === 'object' ? this.properties.displaySubPropertyType : this.properties.displayType) {
      case "bool":
        displayPaneFields.push(...[
          PropertyPaneTextField('displayBoolTrue',{
            label: 'Value when true'
          }),
          PropertyPaneTextField('displayBoolFalse',{
            label: 'Value when false'
          }),
        ]);
        break;
    }

    //Advanced Display Template
    displayPaneFields.push(...[
      PropertyFieldTextWithCallout('displayTemplate', {
        calloutTrigger: CalloutTriggers.Hover,
        key: 'displayTemplate',
        label: 'Template (Advanced)',
        calloutWidth: 200,
        value: this.properties.displayTemplate,
        calloutContent: displayTemplateCallout(),
      }),
      PropertyPaneDropdown('displayUndefined', {
        label: 'When value is undefined',
        options: [
          {key:'blank', text:'Show blank value'},
          {key:'undefined', text:'Show "undefined"'},
          {key:'custom', text:'Use custom value'}
        ]
      })
    ]);

    if(this.properties.displayUndefined === 'custom') {
      displayPaneFields.push(
        PropertyPaneTextField('displayUndefinedCustom', {
          label: 'Custom text for undefined',
        })
      );
    }


    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Source',
              groupFields: [
                PropertyPaneDropdown('sourceId', {
                  label: "Source",
                  options: sourceOptions,
                  selectedKey: this.properties.sourceId,
                }),
                PropertyPaneDropdown('propertyId', {
                  label: "Property",
                  options: propertyOptions,
                  selectedKey: this.properties.propertyId,
                }),
              ]
            },
            {
              groupName: 'Label',
              groupFields: labelPaneFields,
            },
            {
              groupName: 'Display',
              groupFields: displayPaneFields,
            },
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    if (propertyPath === 'sourceId') {
      //New source, so pick the first property (so we can subscribe)
      this.properties.propertyId = this.context.dynamicDataProvider.tryGetSource(this.properties.sourceId).getPropertyDefinitions()[0].id;
    }

    //If either of the subscription props change, we need to unregister and then register
    if (propertyPath === 'sourceId' || propertyPath === 'propertyId') {

      if (this._lastSourceId && this._lastPropertyId) {
        try {
          //previously registered, so unregister
          this.context.dynamicDataProvider.unregisterPropertyChanged(this._lastSourceId, this._lastPropertyId, this.render);
        } catch (e) {
          this.logError("Unable to unregister previous dynamicDataProvider propertyChanged - likely the source was removed", e);
        }
      }

      //Get the title of the current property
      const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(this.properties.sourceId);
      if(source) {
        let props = source.getPropertyDefinitions().filter(prop => {
          return prop.id == this.properties.propertyId;
        });
        if(props.length) {
          this.properties.propertyTitle = props[0].title;
        }
      }

      //If the label isn't set or it equals the previous property title, go ahead and change it
      // This makes the label automatic but keeps any customizations
      if(typeof this.properties.labelText == "undefined" || this.properties.labelText == this._lastPropertyTitle) {
        this.properties.labelText = this.properties.propertyTitle;
      }

      //register for changes (calls render whenever there is a change)
      this.context.dynamicDataProvider.registerPropertyChanged(this.properties.sourceId, this.properties.propertyId, this.render);
      this._lastSourceId = this.properties.sourceId;
      this._lastPropertyId = this.properties.propertyId;
      this._lastPropertyTitle = this.properties.propertyTitle;
      this._registrationComplete = true;
    }
  }
}
