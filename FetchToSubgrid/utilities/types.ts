export interface EntityAttribute {
  linkEntityAlias: string | undefined;
  name: string;
  attributeAlias: string;
}

export type Dictionary<T> = { [key: string]: T };

export type Entity = ComponentFramework.WebApi.Entity;

export type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;

export type RetriveRecords = ComponentFramework.WebApi.RetrieveMultipleResponse;

export type ItemProps = {
  timeZoneDefinitions: any,
  item: Entity;
  isLinkEntity: boolean;
  entityMetadata: EntityMetadata;
  attributeType: number;
  fieldName: string;
  entity: Entity;
  fetchXml: string | null;
  index: number;
};
