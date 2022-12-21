import { IColumn } from '@fluentui/react';
import { IInputs } from '../generated/ManifestTypes';
import { getEntityName, getLinkEntitiesNames,
  getAttributesFieldNames } from '../Utilities/utilities';

let _context: ComponentFramework.Context<IInputs>;
type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;
type RetriveRecords = ComponentFramework.WebApi.RetrieveMultipleResponse;

export const setContext = (context: ComponentFramework.Context<IInputs>) => {
  _context = context;
};

export const getPagingLimit = (): number =>
// @ts-ignore
  _context.userSettings.pagingLimit
;

export const getRecordsCount = async (fetchXml: string): Promise<number> => {
  const entityName = getEntityName(fetchXml);
  const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
  const records = await _context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);
  return records.entities.length;
};

export const getEntityMetadata = async (entityName: string, attributesFieldNames: string[]):
 Promise<EntityMetadata> => {
  const entityMetadata = await _context.utils.getEntityMetadata(entityName, attributesFieldNames);
  return entityMetadata;
};

export const getColumns = async (fetchXml: string | null): Promise<IColumn[]> => {
  const attributesFieldNames: string[] = getAttributesFieldNames(fetchXml ?? '');
  const entityName = getEntityName(fetchXml ?? '');

  const entityMetadata = await getEntityMetadata(entityName, attributesFieldNames);
  const displayNameCollection: { [entityName: string]: EntityMetadata } =
   entityMetadata.Attributes._collection;

  const linkEntityNameAndAttributes = getLinkEntitiesNames(fetchXml ?? '');
  const columns: IColumn[] = [];

  const entityNames: Array<string> = Object.keys(linkEntityNameAndAttributes);
  const entityFieldNames: string[][] = Object.values(linkEntityNameAndAttributes);

  const data: { [entityName: string]: EntityMetadata } = {};

  for (const [i, fieldNames] of Array.from(entityFieldNames.entries())) {
    const attributeNames: string[] = fieldNames.slice(1);
    const medData = await getEntityMetadata(entityNames[i], attributeNames);
    const displayNameCollection: EntityMetadata = medData.Attributes._collection;
    data[entityNames[i]] = displayNameCollection;
  }

  entityFieldNames.forEach((attr: any, index: number) => {
    const entityName: string = entityNames[index];

    for (let i = 1; i < attr.length; i++) {
      const alias: string = attr[0];
      const attributeName: string = attr[i];
      const display: string = data[entityName][attributeName].DisplayName;
      columns.push({
        name: `${display} (${alias})`,
        fieldName: `${alias}.${attributeName}`,
        key: `col-el-${i}`,
        minWidth: 10,
        isResizable: true,
        isMultiline: false,
      });
    }
  });

  attributesFieldNames.forEach((name, index) => {
    columns.push({
      name: displayNameCollection[name].DisplayName,
      fieldName: name,
      key: `col-${index}`,
      minWidth: 10,
      isResizable: true,
      isMultiline: false,
    });
  });

  return columns;
};

export const openLookupForm = (entity: any, fieldName: string): void => {
  _context.navigation.openForm(
    {
      entityName: entity[`_${fieldName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`],
      entityId: entity[`_${fieldName}_value`],
    },
  );
};

export const openLinkEntityRecord = (entity: any, fieldName: string): void => {
  const entityName = entity[`${fieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`];
  _context.navigation.openForm(
    {
      entityName,
      entityId: entity[fieldName],
    },
  );
};

export const openPrimaryEntityForm = (entity: any, entityName: string): void => {
  _context.navigation.openForm(
    {
      entityName,
      entityId: entity[`${entityName}id`],
    },
  );
};

export const openRecord = (entityName: string, entityId: string): void => {
  _context.navigation.openForm(
    {
      entityName,
      entityId,
    },
  );
};

export const getRecords = async (fetchXml: string | null): Promise<RetriveRecords> => {
  const entityName = getEntityName(fetchXml ?? '');
  const encodeFetchXml = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
  return await _context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
};
