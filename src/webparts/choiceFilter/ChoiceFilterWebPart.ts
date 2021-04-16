import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Log } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart,  } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle, PropertyPaneDropdown, PropertyPaneHorizontalRule, PropertyPaneLabel, PropertyPaneButton, IPropertyPaneGroup, PropertyPaneCheckbox } from "@microsoft/sp-property-pane";
import { IDynamicDataCallables, IDynamicDataPropertyDefinition, IDynamicDataSource } from '@microsoft/sp-dynamic-data';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'ChoiceFilterWebPartStrings';
import { IChoiceFilterProps, ChoiceFilter } from './components/ChoiceFilter';
import { sp } from "@pnp/sp";
import {IListInfo, IViewInfo, IFieldInfo} from './IListInfo';
import { uniq } from '@microsoft/sp-lodash-subset';

export interface IChoiceFilterWebPartProps {
  title: string;
  labelShow: boolean;
  labelText: string;
  choiceType: string;
  customChoices: any[];
  listId: string;
  viewId: string;
  viewQuery: string;
  viewFields: string[];
  keyField: string;
  textField: string;
  allowNone: boolean;
  cascadingFilters: any[];
  syncQS: boolean;
  QSkey: string;
  sendAsArray: boolean;
}

export default class ChoiceFilterWebPart extends BaseClientSideWebPart<IChoiceFilterWebPartProps> implements IDynamicDataCallables {

  private _selectedKey?: any;
  private _selectedText?: string;

  private _listInfo: Map<string,IListInfo>;
  private _loadingListChoices: boolean;
  private _listResults: Array<any>;

  private _previousQSkey: string;

  public onInit(): Promise<void> {
    this.render = this.render.bind(this);

    this.context.dynamicDataSourceManager.initializeSource(this);
    this._listInfo = new Map<string,IListInfo>();

    this._previousQSkey = this.properties.QSkey;

    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      if(qsParams.has(this.properties.QSkey)) {
        this._selectedKey = qsParams.get(this.properties.QSkey);
      }
    }

    //Go get list choices
    if(this.properties.choiceType === 'list') {
      this._loadingListChoices = true;
      this.getListChoices().then(results => {
        this._listResults = results;
        this._loadingListChoices = false;
        this.render();
      }).catch(e => {
        this.logError('Failed during initial loading of list choices', e);
      });
    }

	  return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context,
        defaultCachingTimeoutSeconds: 60,
      });
    });

  }

  private logError(message:string, e:any) {
    Log.error(message,e,this.context.serviceScope);
  }

  public render(): void {
    const needsConfiguration: boolean = (this.properties.choiceType == "list" &&
      !(this.properties.listId && this.properties.viewId && this.properties.viewQuery &&
        this.properties.viewFields && this.properties.keyField && this.properties.textField));

        if (this.renderedOnce === false && !needsConfiguration) {
          if(this.properties.choiceType == 'list') {
            this.properties.cascadingFilters.forEach(filter => {
              this.context.dynamicDataProvider.registerPropertyChanged(filter.source, filter.prop, this.render);
            });
          }
        }

    let options: IDropdownOption[] = this.getOptions(needsConfiguration);

    if(!(this.properties.choiceType === 'list' && this._loadingListChoices)) {
      if(typeof this._selectedKey !== "undefined") {
        let selectedOption = options.filter(opt => {
          return opt.key == this._selectedKey;
        });
        if (!selectedOption.length) {
          //Selection not found
          this._selectedKey = undefined;
          this._selectedText = undefined;
          if(this.properties.allowNone || !options.length) {
            //if none is allowed, or there aren't any options - tell people about it
            this.context.dynamicDataSourceManager.notifyPropertyChanged('filterKey');
            this.context.dynamicDataSourceManager.notifyPropertyChanged('filterText');
            this.syncQueryString();
          }
        }
      }
      if(!this.properties.allowNone && typeof this._selectedKey == "undefined" && options.length) {
        this._selectedKey = options[0].key;
        this._selectedText = options[0].text;
        this.context.dynamicDataSourceManager.notifyPropertyChanged('filterKey');
        this.context.dynamicDataSourceManager.notifyPropertyChanged('filterText');
        this.syncQueryString();
      }
    }

    const element: React.ReactElement<IChoiceFilterProps > = React.createElement(
      ChoiceFilter,
      {
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
        isLoading: this.properties.choiceType === 'list' && this._loadingListChoices,
        options: options,
        selectedKey: this._selectedKey,
        allowNone: this.properties.allowNone,
        onChange: (option:IDropdownOption) => {
          this._selectedKey = option.key;
          this._selectedText = option.text;
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterKey');
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterText');
          this.syncQueryString();
          this.render();
        },
        onBlank: () => {
          this._selectedKey = undefined;
          this._selectedText = undefined;
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterKey');
          this.context.dynamicDataSourceManager.notifyPropertyChanged('filterText');
          this.syncQueryString();
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private syncQueryString(): void {
    if(this.properties.syncQS && this.properties.QSkey) {
      const qsParams = new URLSearchParams(location.search);
      if(typeof this._selectedKey == "undefined" && qsParams.has(this.properties.QSkey)) {
        qsParams.delete(this.properties.QSkey);
      } else {
        qsParams.set(this.properties.QSkey, this._selectedKey);
      }
      if (this.properties.QSkey !== this._previousQSkey && this._previousQSkey) {
        if (qsParams.has(this._previousQSkey)) {
          qsParams.delete(this._previousQSkey);
        }
        this._previousQSkey = this.properties.QSkey;
      }
      window.history.replaceState({},'',`${location.pathname}?${qsParams}`);
    }
  }

  private getResultValue(fieldName:string, result:any): any {
    if(fieldName.indexOf('/') > 0){
      const props = fieldName.split('/');
      return result[props[0]][0][props[1]];
    } else {
      return result[fieldName];
    }
  }

  private evaluateFilter(resultValue: any, operation: string, propValue: any): boolean {
    if(typeof resultValue == "undefined" && typeof propValue == "undefined" && operation == "eq") {
      return true;
    }
    if(typeof resultValue == "undefined" || typeof propValue == "undefined") {
      return false;
    }

    switch(operation) {
      case "eq":
        return resultValue == propValue;
      case "gt":
        return resultValue > propValue;
      case "gte":
        return resultValue >= propValue;
      case "lt":
        return resultValue < propValue;
      case "lte":
        return resultValue <= propValue;
      case "contains":
        return resultValue.toString().toLowerCase().indexOf(propValue.toString().toLowerCase()) >= 0;
      case "starts":
        return resultValue.toString().toLowerCase().startsWith(propValue.toString().toLowerCase());
      case "ends":
        return resultValue.toString().toLowerCase().substr(resultValue.toString().length - propValue.toString().length) == propValue.toString().toLowerCase();
      case "containsCS":
        return resultValue.toString().indexOf(propValue.toString()) >= 0;
      case "startsCS":
        return resultValue.toString().startsWith(propValue.toString());
      case "endsCS":
        return resultValue.toString().substr(resultValue.toString().length - propValue.toString().length) == propValue.toString();
      default:
        return true;
    }
  }

  private filterResults(results: Array<any>): Array<any> {
    return results
      .filter(result => {
        let isIncluded: boolean = true;
        this.properties.cascadingFilters.forEach(filter => {
          //Assumes all filters are AND
          if(!isIncluded) {
            return; //skip eval if already false
          }
          let propValue: any;
          const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(filter.source);
          if(!source) {
            this.logError(`Unable to connect to the data source (${filter.source})`, new Error('datasource not found'));
            return; //filter not applied
          }
          try {
            propValue = source.getPropertyValue(filter.prop);
            if(typeof propValue !== "undefined" && filter.useSub) {
              propValue = propValue[filter.sub];
            }
          }
          catch(e) {
            this.logError(`An error has occurred while retrieving the property value (${filter.prop}).`, e);
            return; //filter not applied
          }
          isIncluded = isIncluded && this.evaluateFilter(this.getResultValue(filter.field, result), filter.operation, propValue);
        });
        return isIncluded;
      });
  }

  private getOptions(needsConfiguration:boolean): IDropdownOption[] {
    let options: Array<IDropdownOption> = [];

    if(this.properties.choiceType === 'custom') {
      //TODO add filtering stuff
      options.push(
        ...this.filterResults(this.properties.customChoices)
        .map(result => {
          return {
            key: result.key,
            text: result.text,
          };
        })
      );
    } else {
      if(typeof this._listResults !== "undefined" && !this._loadingListChoices) {
        options.push(
          ...this.filterResults(this._listResults)
            .map(result => {
              return {
                key: this.getResultValue(this.properties.keyField, result),
                text: this.getResultValue(this.properties.textField, result).replace(/&#39;/g, "'")
                //Above needs to be evaluated for smarter (more complete) escaping
              };
        }));
      } else if(!this._loadingListChoices && !needsConfiguration) {
        //Go get list choices
        this._loadingListChoices = true;
        this.getListChoices().then(results => {
          this._listResults = results;
          this._loadingListChoices = false;
          this.render();
        }).catch(e => {
          this.logError('Failed to load list choices', e);
        });
      }
    }

    return options;
  }

  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'filterKey',
        title: 'Filter Key',
        description: 'The key value of the selected filter choice',

      },
      {
        id: 'filterText',
        title: 'Filter Text',
        description: 'The display text of the selected filter choice',
      },
    ];
  }

  public getPropertyValue(propertyId: string): any {
    switch(propertyId) {
      case 'filterKey':
        return this.properties.sendAsArray ? [{value:this._selectedKey}] : this._selectedKey;
      case 'filterText':
        return this.properties.sendAsArray ? [{value:this._selectedText}] : this._selectedText;
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
    let propsSetBehindTheScenes: boolean = false;
    let filterPaneFields = new Array<any>();

    let choicesPaneFields: Array<any> = [
      PropertyPaneDropdown('choiceType', {
        label: "Source",
        options: [
          {key:'custom', text:'Custom'},
          {key:'list', text:'List/Library'},
        ]
      })
    ];

    if(this.properties.choiceType === 'custom') {
      choicesPaneFields.push(
        PropertyFieldCollectionData('customChoices', {
          key: 'customChoices',
          label: 'Custom Choices',
          panelHeader: 'Custom Choices',
          manageBtnLabel: 'Manage Choices',
          value: this.properties.customChoices,
          fields: [
            {
              id: 'key',
              title: 'Key',
              type: CustomCollectionFieldType.string,
              required: true
            },
            {
              id: 'text',
              title: 'Text',
              type: CustomCollectionFieldType.string,
              required: true
            },
            {
              id: 'filterValue',
              title: 'Filterable Value',
              type: CustomCollectionFieldType.string,
            }
          ]
        })
      );
      filterPaneFields.push(...this.buildFilterPaneFields([
        {key:'key', text:'Key'},
        {key:'text', text:'Text'},
        {key:'filterValue', text:'Filterable Value'},
      ], 'filterValue'));
    }

    if(this.properties.choiceType === 'list') {

      choicesPaneFields.push(
        PropertyFieldListPicker('listId', {
          label: 'List/Library',
          selectedList: this.properties.listId,
          includeHidden: false,
          orderBy: PropertyFieldListPickerOrderBy.Title,
          onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
          properties: this.properties,
          context: this.context,
          onGetErrorMessage: null,
          deferredValidationTime: 0,
          key: 'listId',
        })
      );

      //List Details
      if(this.properties.listId) {
        //If a list has been chosen, we need view details

        if(!this._listInfo.has(this.properties.listId)) {
          //No view details yet, so go get them
          choicesPaneFields.push(
            PropertyPaneLabel('',{
              text: 'Loading View Details...',
            })
          );
          this.getListInfo(this.properties.listId)
            .then((listinfo:IListInfo) => {
              this.context.propertyPane.refresh();
            })
            .catch(e => {
              this.logError('Failed to get List Info for ' + this.properties.listId, e);
            });
        } else {
          //View Dropdown
          const listinfo = this._listInfo.get(this.properties.listId);

          const viewOptions: IDropdownOption[] = [];
          listinfo.views.forEach((viewInfo:IViewInfo, id:string) => {
            viewOptions.push({key: id, text: viewInfo.title});
          });
          if(!this.properties.viewId && viewOptions.length) {
            this.properties.viewId = viewOptions[0].key.toString();
            const viewinfo = listinfo.views.get(this.properties.viewId);
            this.properties.viewQuery = viewinfo.query;
            this.properties.viewFields = viewinfo.viewFields;
            propsSetBehindTheScenes = true;
          }
          choicesPaneFields.push(
            PropertyPaneDropdown('viewId', {
              label: 'View',
              options: viewOptions.sort((a,b) => {
                if(a.text < b.text) return -1;
                if(a.text > b.text) return 1;
                return 0;
              }),
            })
          );

          if(this.properties.viewId) {
            //Field Dropdowns
            const viewinfo = listinfo.views.get(this.properties.viewId);

            let fieldOptions: IDropdownOption[] = [];
            viewinfo.fieldChoices.forEach((field:string) => {
              fieldOptions.push({key: field, text: listinfo.fields.get(field)});
            });
            fieldOptions = fieldOptions.sort((a,b) => {
              if(a.text < b.text) return -1;
              if(a.text > b.text) return 1;
              return 0;
            });

            if(!this.properties.keyField) {
              //If one isn't set, neither are, so set them
              // we know ID will always be included (so at least 1 field)
              this.properties.keyField = 'ID';
              if(viewinfo.fieldChoices.indexOf('Title') > -1) {
                this.properties.textField = 'Title';
              } else {
                this.properties.textField = fieldOptions[0].key.toString();
              }
              propsSetBehindTheScenes = true;
            }

            choicesPaneFields.push(
              PropertyPaneDropdown('keyField', {
                label: 'Key Field',
                options: fieldOptions,
              }),
              PropertyPaneDropdown('textField', {
                label: 'Text Field',
                options: fieldOptions,
              })
            );

            filterPaneFields.push(...this.buildFilterPaneFields(fieldOptions, 'ID'));
          }
        }
      }
    }

    choicesPaneFields.push(
      PropertyPaneToggle('allowNone', {
        label: 'Allow empty selection'
      })
    );

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

    if(propsSetBehindTheScenes) {
      this.render();
    }

    let groups: IPropertyPaneGroup[] = [
      {
        groupName: 'Label',
        groupFields: labelPaneFields,
      },
      {
        groupName: "Choices",
        groupFields: choicesPaneFields,
      },
    ];

    if(filterPaneFields.length) {
      groups.push(
        {
          groupName: "Cascading Filters",
          groupFields: filterPaneFields,
        }
      );
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
    qsPaneFields.push(
      PropertyPaneToggle('sendAsArray',{
        label: 'Send as array'
      })
    );

    groups.push(
      {
        groupName: 'Advanced',
        groupFields: qsPaneFields,
      }
    );

    return {
      pages: [
        {
          groups: groups
        }
      ]
    };
  }

  private buildFilterPaneFields(fieldOptions: IDropdownOption[], defaultField: string): any[] {
    const filterFields = [];

    const sourceOptions: IDropdownOption[] = this.context.dynamicDataProvider.getAvailableSources()
      .filter(source => {
        return source.id !== this.context.dynamicDataSourceManager.sourceId;
      })
      .map(source => {
        return {
          key: source.id,
          text: source.metadata.title,
        };
      });

    this.properties.cascadingFilters.forEach((filter:any, index:number) => {
      const filterName: string = 'Filter ' + (index+1).toString();

      let propertyOptions: IDropdownOption[] = [];
      const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(filter.source);
      if(source) {
        propertyOptions = source.getPropertyDefinitions().map(prop => {
          return {
            key: prop.id,
            text: prop.title,
          };
        });
      }

      filterFields.push(
        PropertyPaneDropdown(`listFilter_${index}_field`, {
          label: filterName + ' field',
          options: fieldOptions,
          selectedKey: filter.field,
        }),
        PropertyPaneDropdown(`listFilter_${index}_operation`, {
          label: filterName + ' operation',
          options: [
            {key:'eq', text:"="},
            {key:'gt', text:'>'},
            {key:'gte', text:'>='},
            {key:'lt', text:'<'},
            {key:'lte', text:'<='},
            {key:'contains', text:'contains'},
            {key:'starts', text:'starts with'},
            {key:'ends', text:'ends with'},
            {key:'containsCS', text:'contains (case sensitive)'},
            {key:'startsCS', text:'starts with (case sensitive)'},
            {key:'endsCS', text:'ends with (case sensitive)'},
          ],
          selectedKey: filter.operation,
        }),
        PropertyPaneDropdown(`listFilter_${index}_source`, {
          label: filterName + ' source',
          options: sourceOptions,
          selectedKey: filter.source,
        }),
        PropertyPaneDropdown(`listFilter_${index}_prop`, {
          label: filterName + ' property',
          options: propertyOptions,
          selectedKey: filter.prop
        }),
        PropertyPaneCheckbox(`listFilter_${index}_useSub`, {
          text: 'Use sub property',
          checked: filter.useSub,
        })
      );

      if(filter.useSub) {
        let keyOptions: IDropdownOption[] = [];
        if(source) {
          try {
            let sample: any = source.getPropertyValue(filter.prop);
            keyOptions = Object.keys(sample).map(key => {
              return {key:key, text:key};
            });
          }
          catch(e) {
            this.logError('Unable to get sub properties from source property', e);
          }
        }
        let sub:string = filter.sub;
          //If the current displayObjectProperty isn't in the list, then pick the first one
        if(typeof filter.sub == "undefined" || !keyOptions.filter((opt:IDropdownOption) => {
          return opt.key.toString() === filter.sub;
        }).length) {
          this.properties.cascadingFilters[index].sub = keyOptions[0].key.toString();
          sub = keyOptions[0].key.toString();
          this.render();
        }
        filterFields.push(
          PropertyPaneDropdown(`listFilter_${index}_sub`, {
            label: 'Sub property',
            options: keyOptions,
            selectedKey: sub,
          })
        );
      }

      filterFields.push(
        PropertyPaneButton(`listFilter_${index}_remove`, {
          text: 'Remove ' + filterName,
          icon: 'Delete',
          onClick: () => {
            try{
              this.context.dynamicDataProvider.unregisterPropertyChanged(this.properties.cascadingFilters[index].source, this.properties.cascadingFilters[index].prop, this.render);
            }
            catch(e){
              this.logError('Error unregistering filter subscription', e);
            }
            this.properties.cascadingFilters.splice(index,1,...this.properties.cascadingFilters.slice(index+1));
          }
        }),
        PropertyPaneHorizontalRule()
      );
    });

    return [
      ...filterFields,
      PropertyPaneButton('addFilter', {
        text: 'Add Filter',
        icon: 'CirclePlus',
        onClick: () => {
          let defaultSource: string = sourceOptions[0].key.toString();
          let defaultProp: string;
          const source: IDynamicDataSource = this.context.dynamicDataProvider.tryGetSource(defaultSource);
          if(source) {
            defaultProp = source.getPropertyDefinitions()[0].id;
          }
          this.properties.cascadingFilters.push({
            field: defaultField,
            operation: 'eq',
            source: defaultSource,
            prop: defaultProp,
            useSub: false,
            sub: '',
          });
          this.context.dynamicDataProvider.registerPropertyChanged(defaultSource, defaultProp, this.render);
        },
      })
    ];
  }

  private getListInfo(listId:string): Promise<IListInfo> {
    return new Promise<IListInfo>((resolve: (info: IListInfo) => void, reject: (error:any) => void): void => {
      if(this._listInfo.has(listId)) {
        resolve(this._listInfo.get(listId));
      } else {
        sp.web.lists.getById(this.properties.listId).select('Id','Title','Views/Id','Views/Title','Views/Hidden','Views/ViewFields','Views/ViewQuery','Fields/InternalName','Fields/Title','Fields/Hidden','Fields/TypeAsString','Fields/IsDependentLookup','Fields/LookupField','Fields/DependentLookupInternalNames').expand('Views','Views/ViewFields','Fields').usingCaching().get()
          .then((data:any) => {
            //Setup list info object (reference for views & fields in pane)
            const listinfo:IListInfo = {
              title: data.Title,
              id: data.Id,
              fields: new Map<string, string>(),
              views: new Map<string, IViewInfo>(),
            };

            //Collect information about non-hidden fields (for reference by views)
            let fields = new Map<string, IFieldInfo>(); //InternalName is key
            data.Fields.filter(fieldEntry => {
              return !fieldEntry.Hidden;
            }).forEach(fieldEntry => {
              fields.set(fieldEntry.InternalName, {
                internalName: fieldEntry.InternalName,
                title: fieldEntry.Title,
                type: fieldEntry.TypeAsString,
                lookupField: fieldEntry.LookupField,
                dependentLookup: fieldEntry.IsDependentLookup
              });
            });

            //Process non-hidden views (and their fields)
            listinfo.fields.set('ID','ID');
            listinfo.fields.set('Title','Title');
            data.Views.filter(viewEntry => {
              return !viewEntry.Hidden;
            }).forEach(viewEntry => {
              const viewFields = ['ID']; //ensure ID is always included
              const fieldOptions = ['ID'];
              viewEntry.ViewFields.Items.forEach(viewfield => {
                const field = fields.get(viewfield);
                switch(field.type) {
                  case 'Lookup':
                    viewFields.push(viewfield);
                    if(field.dependentLookup) {
                      //treat it like normal
                      fieldOptions.push(viewfield);
                      listinfo.fields.set(viewfield,field.title);
                    } else {
                      let internalNameId = `${viewfield}/lookupId`;
                      fieldOptions.push(internalNameId);
                      listinfo.fields.set(internalNameId,`${field.title} (Id)`);
                      if(field.lookupField !== 'ID') {
                        //If the primary lookupfield is not ID (covered above) we get that too
                        let internalName = `${viewfield}/lookupValue`;
                        fieldOptions.push(internalName);
                        listinfo.fields.set(internalName, `${field.title} (${field.lookupField})`);
                      }
                    }
                    break;
                  case 'User':
                    viewFields.push(viewfield);
                    //Provide the option to pull back specific person fields
                    let inId = `${viewfield}/id`;
                    listinfo.fields.set(inId, `${field.title} (Id)`);
                    let inTitle = `${viewfield}/title`;
                    listinfo.fields.set(inTitle, `${field.title} (Name)`);
                    let inE = `${viewfield}/email`;
                    listinfo.fields.set(inE, `${field.title} (Email)`);
                    let inP = `${viewfield}/picture`;
                    listinfo.fields.set(inP, `${field.title} (Picture URL)`);
                    let inUN = `${viewfield}/sip`;
                    listinfo.fields.set(inUN, `${field.title} (SIP)`);
                    fieldOptions.push(inId, inTitle, inE, inP, inUN);
                    break;
                  case 'URL':
                    viewFields.push(viewfield);
                    fieldOptions.push(viewfield);
                    listinfo.fields.set(viewfield, `${field.title} (URL)`);
                    let inDesc = `${viewfield}.desc`;
                    fieldOptions.push(inDesc);
                    listinfo.fields.set(inDesc, `${field.title} (Description)`);
                    break;
                  case 'UserMulti':
                  case 'LookupMulti':
                  case 'TaxonomFieldTypeMulti':
                  case 'MultiChoice':
                  case 'Attachments':
                    //Unsupported types
                    break;
                  default:
                    if(viewfield == 'LinkTitle' || viewfield == 'LinkTitleNoMenu' || viewfield == 'LinkTitle2') {
                      viewFields.push('Title');
                      fieldOptions.push('Title');
                    } else {
                      fieldOptions.push(viewfield);
                      viewFields.push(viewfield);
                      listinfo.fields.set(viewfield,field.title);
                    }
                }
              });

              listinfo.views.set(viewEntry.Id, {
                id: viewEntry.Id,
                title: viewEntry.Title,
                query: viewEntry.ViewQuery,
                viewFields: uniq(viewFields),
                fieldChoices: uniq(fieldOptions),
              });
            });

            this._listInfo.set(listinfo.id, listinfo);
            resolve(listinfo);
        }).catch(e => {
          reject(e);
        });
      }
    });
  }

  private getListChoices(): Promise<any> {
    return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error:any) => void): void => {
      const viewXml = `<View><Query>${this.properties.viewQuery}</Query><ViewFields>${this.fieldRefs(this.properties.viewFields)}</ViewFields></View>`;
      sp.web.lists.getById(this.properties.listId).renderListData(viewXml)
        .then(results => {
          resolve(results.Row);
        }).catch(e => {
          reject(e);
        });
    });
  }

  private fieldRefs(fields:string[]): string {
    return fields.map(field => {
      return `<FieldRef Name="${field}"/>`;
    }).join('');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    if (propertyPath === 'listId') {
      this.properties.viewId = undefined;
      this.properties.viewQuery = undefined;
      this.properties.viewFields = undefined;
      this.properties.keyField = undefined;
      this.properties.textField = undefined;
      this._listResults = undefined;
    }
    if (propertyPath === 'viewId') {
      if(this.properties.listId && this.properties.viewId &&
         this._listInfo.has(this.properties.listId)) {
        let listinfo = this._listInfo.get(this.properties.listId);
        if(listinfo.views.has(this.properties.viewId)) {
          let viewinfo = listinfo.views.get(this.properties.viewId);
          this.properties.viewQuery = viewinfo.query;
          this.properties.viewFields = viewinfo.viewFields;
        }
      }
      this.properties.keyField = undefined;
      this.properties.textField = undefined;
      this._listResults = undefined;
    }

    if (propertyPath === 'choiceType' || propertyPath === 'listId' || propertyPath === 'viewId') {
      this.properties.cascadingFilters.forEach(filter => {
        try {
          //unregister dynamic data listener
          this.context.dynamicDataProvider.unregisterPropertyChanged(filter.source, filter.prop, this.render);
        }
        catch (e) {
          this.logError('Error unregistering filter subscription', e);
        }
      });
      this.properties.cascadingFilters = [];
    }

    if (propertyPath.startsWith('listFilter')) {
      //Set the dynamic property value in collection, then remove it
      const pathPart = propertyPath.split('_');
      let propIndex = parseInt(pathPart[1]);
      if(this.properties.cascadingFilters.length > propIndex) {
        if(pathPart[2] === 'source' || pathPart[2] === 'prop') {
          try {
            //unregister dynamic data listener
            this.context.dynamicDataProvider.unregisterPropertyChanged(this.properties.cascadingFilters[propIndex].source, this.properties.cascadingFilters[propIndex].prop, this.render);
          }
          catch (e) {
            this.logError('Error unregistering filter subscription', e);
          }
        }
        this.properties.cascadingFilters[propIndex][pathPart[2]] = this.properties[propertyPath];
        if(pathPart[2] === 'source' || pathPart[2] === 'prop') {
          //register dynamic data Listener
          this.context.dynamicDataProvider.registerPropertyChanged(this.properties.cascadingFilters[propIndex].source, this.properties.cascadingFilters[propIndex].prop, this.render);
        }
      }
      this.properties[propertyPath] = undefined;
    }

    if (propertyPath === 'syncQS' || propertyPath === 'QSkey') {
      this.syncQueryString();
    }
  }
}
