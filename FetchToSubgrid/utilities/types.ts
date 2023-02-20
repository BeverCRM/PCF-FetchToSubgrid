export type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;

export type RetriveRecords = ComponentFramework.WebApi.RetrieveMultipleResponse;

export type Entity = ComponentFramework.WebApi.Entity;

export interface EntityAttribute {
  linkEntityAlias?: string;
  name: string;
  attributeAlias: string;
}

export type IItemProps = {
  timeZoneDefinitions: any;
  item: Entity;
  isLinkEntity: boolean;
  entityMetadata: EntityMetadata;
  attributeType: number;
  fieldName: string;
  entity: Entity;
  fetchXml: string | null;
  index: number;
};

export interface IFetchSubgridProps {
  fetchXml: string | null;
  defaultPageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
}

export type Dictionary<T> = { [key: string]: T }
