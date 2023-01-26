export type Dictionary = { [key: string]: string[] };

export type Entity = ComponentFramework.WebApi.Entity;

export type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;

export type EntityMetadataDictionary = { [entityName: string]: EntityMetadata };

export type RetriveRecords = ComponentFramework.WebApi.RetrieveMultipleResponse;

export type IItemProps = {
  item: Entity;
  isLinkEntity: boolean;
  entityMetadata: EntityMetadata;
  attributeType: number;
  fieldName: string;
  entity: Entity;
  fetchXml: string | null;
  index: number;
};
