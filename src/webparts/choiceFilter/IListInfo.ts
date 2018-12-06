export interface IFieldInfo {
  internalName: string;
  title: string;
  type: string;
  lookupField: string;
  dependentLookup: boolean;
}

export interface IViewInfo {
  title: string;
  id: string;
  viewFields: string[]; //for query
  fieldChoices: string[]; //for property pane dropdown
  query: string;
}

export interface IListInfo {
  title: string;
  id: string;
  fields: Map<string, string>; //<InternalName,Title>
  views: Map<string, IViewInfo>; //<Id,ViewInfo>
}
