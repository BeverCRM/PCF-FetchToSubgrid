export type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;

export type RetriveRecords = ComponentFramework.WebApi.RetrieveMultipleResponse;

export type Entity = ComponentFramework.WebApi.Entity;

export type DialogResponse = ComponentFramework.NavigationApi.ConfirmDialogResponse

export type Dictionary<T> = { [key: string]: T };

export interface IService<T> {
  _service: T;
}

export interface OrderInFetchXml {
  [attribute: string]: boolean;
  isLinkEntity: boolean;
}

export interface EntityAttribute {
  linkEntityAlias?: string;
  name: string;
  attributeAlias: string;
}
