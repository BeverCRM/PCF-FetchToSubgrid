import { IColumn } from '@fluentui/react';
import { IInputs } from '../generated/ManifestTypes';
import {
  getEntityName,
  getLinkEntitiesNames,
  getAttributesFieldNames,
  isAggregate,
  getAliasNames,
} from '../utilities/utilities';

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
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

  fetch.removeAttribute('count');
  fetch.removeAttribute('page');

  const fetchWithoutCount: string = new XMLSerializer().serializeToString(xmlDoc);
  const entityName: string = getEntityName(fetchWithoutCount);

  const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchWithoutCount ?? '')}`;
  const records: RetriveRecords =
    await _context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);

  return records.entities.length;
};

export const getEntityMetadata = async (entityName: string, attributesFieldNames: string[]):
  Promise<EntityMetadata> => {
  const entityMetadata: EntityMetadata = await _context.utils.getEntityMetadata(
    entityName,
    [...attributesFieldNames]);

  return entityMetadata;
};

export const getColumns = async (fetchXml: string | null): Promise<IColumn[]> => {
  const attributesFieldNames: string[] = getAttributesFieldNames(fetchXml ?? '');
  const entityName: string = getEntityName(fetchXml ?? '');

  const entityMetadata: EntityMetadata = await getEntityMetadata(entityName, attributesFieldNames);
  const displayNameCollection: { [entityName: string]: EntityMetadata } =
   entityMetadata.Attributes._collection;

  const linkEntityNameAndAttributes: {[key: string]: string[]} =
   getLinkEntitiesNames(fetchXml ?? '');
  const columns: IColumn[] = [];

  const entityNames: Array<string> = Object.keys(linkEntityNameAndAttributes);
  const entityFieldNames: string[][] = Object.values(linkEntityNameAndAttributes);

  const data: { [entityName: string]: EntityMetadata } = {};

  for (const [i, fieldNames] of Array.from(entityFieldNames.entries())) {
    const attributeNames: string[] = fieldNames.slice(1);
    const mataData: EntityMetadata = await getEntityMetadata(entityNames[i], attributeNames);
    const displayNameCollection: EntityMetadata = mataData.Attributes._collection;
    data[entityNames[i]] = displayNameCollection;
  }

  const hasAggregate: boolean = isAggregate(fetchXml ?? '');
  const aggregateAttrNames: string[] | null = hasAggregate ? getAliasNames(fetchXml ?? '') : null;

  attributesFieldNames.forEach((name, index) => {
    let displayName = displayNameCollection[name].DisplayName;

    if (aggregateAttrNames?.length) {
      displayName = aggregateAttrNames[index];
      name = aggregateAttrNames[index];
    }

    columns.push({
      name: displayName,
      fieldName: name,
      key: `col-${index}`,
      minWidth: 10,
      isResizable: true,
      isMultiline: false,
    });
  });

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

  return columns;
};

export const getRecords = async (fetchXml: string | null): Promise<RetriveRecords> => {
  const entityName: string = getEntityName(fetchXml ?? '');
  const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
  return await _context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
};

export const openRecord = (entityName: string, entityId: string): void => {
  _context.navigation.openForm({
    entityName,
    entityId,
  });
};

export const openLookupForm = (
  entity: ComponentFramework.WebApi.Entity,
  fieldName: string): void => {
  const entityName: string = entity[`_${fieldName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`];
  const entityId: string = entity[`_${fieldName}_value`];
  openRecord(entityName, entityId);
};

export const openLinkEntityRecord = (
  entity: ComponentFramework.WebApi.Entity,
  fieldName: string): void => {
  const entityName: string = entity[`${fieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`];
  openRecord(entityName, entity[fieldName]);
};

export const openPrimaryEntityForm = (
  entity: ComponentFramework.WebApi.Entity,
  entityName: string): void => {
  openRecord(entityName, entity[`${entityName}id`]);
};
