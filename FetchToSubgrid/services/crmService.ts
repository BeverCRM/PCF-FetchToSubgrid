import { IColumn } from '@fluentui/react';
import { IInputs } from '../generated/ManifestTypes';
import {
  Dictionary,
  RetriveRecords,
  Entity,
  EntityMetadata,
  EntityAttribute,
} from '../utilities/types';
import {
  getEntityName,
  getLinkEntitiesNames,
  getAttributesFieldNames,
  isAggregate,
  getAliasNames,
} from '../utilities/utilities';

let _context: ComponentFramework.Context<IInputs>;

export const setContext = (context: ComponentFramework.Context<IInputs>) => {
  _context = context;
};

// @ts-ignore
export const getPagingLimit = (): number => _context.userSettings.pagingLimit;

export const getTimeZoneDefinitions = async () => {
  // @ts-ignore
  const contextPage = _context.page;

  const request = await fetch(`${contextPage.getClientUrl()}/api/data/v9.0/timezonedefinitions`);
  const results = await request.json();

  return results;
};

export const getWholeNumberFieldName = (
  format: string,
  entity: Entity,
  fieldName: string,
  timeZoneDefinitions: any) => {
  let fieldValue: number = entity[fieldName];

  if (format === '3') { return _context.formatting.formatLanguage(fieldValue); }
  if (format === '2') {
    return timeZoneDefinitions.value.find((e: any) =>
      e.timezonecode === Number(fieldValue)).userinterfacename;
  }

  if (fieldValue) {
    let unit: string;
    if (fieldValue < 60) { unit = 'minute'; }
    else if (fieldValue < 1440) {
      fieldValue /= 60;
      unit = 'hour';
    }
    else {
      fieldValue /= 1440;
      unit = 'day';
    }
    return `${fieldValue} ${unit}${fieldValue === 1 ? '' : 's'}`;
  }
};

export const getRecordsCount = async (fetchXml: string): Promise<number> => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

  fetch.removeAttribute('count');
  fetch.removeAttribute('page');

  const fetchWithoutCount: string = new XMLSerializer().serializeToString(xmlDoc);
  const entityName: string = getEntityName(fetchWithoutCount);

  const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchWithoutCount ?? '')}`;
  const records: RetriveRecords = await _context.webAPI.retrieveMultipleRecords(
    `${entityName}`,
    encodeFetchXml);

  return records.entities.length;
};

export const getEntityMetadata = async (
  entityName: string,
  attributesFieldNames: string[]): Promise<EntityMetadata> => {
  const entityMetadata: EntityMetadata = await _context.utils.getEntityMetadata(
    entityName,
    [...attributesFieldNames]);

  return entityMetadata;
};

export const getRecords = async (fetchXml: string | null): Promise<RetriveRecords> => {
  const entityName: string = getEntityName(fetchXml ?? '');
  const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
  return await _context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
};

export const getColumns = async (fetchXml: string | null): Promise<IColumn[]> => {
  let attributesFieldNames: string[] = getAttributesFieldNames(fetchXml ?? '');
  const entityName: string = getEntityName(fetchXml ?? '');

  let isAllAttribute = false;

  if (attributesFieldNames.length === 0 && entityName) {
    const recors = await getRecords(fetchXml ?? '');
    attributesFieldNames = Object.keys(recors.entities[0]);
    isAllAttribute = true;
  }

  const entityMetadata: EntityMetadata = await getEntityMetadata(entityName, attributesFieldNames);
  const displayNameCollection: Dictionary<EntityMetadata> = entityMetadata.Attributes._collection;

  if (isAllAttribute) {
    attributesFieldNames = Object.keys(displayNameCollection);
  }

  const columns: IColumn[] = [];
  // const linkEntityNameAndAttributes: Dictionary<string[]> = getLinkEntitiesNames(
  //   fetchXml ?? '');

  const linkEntityAttFieldNames: Dictionary<EntityAttribute[]> = getLinkEntitiesNames(
    fetchXml ?? '');
  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes: EntityAttribute[][] = Object.values(
    linkEntityAttFieldNames);

  const hasAggregate: boolean = isAggregate(fetchXml ?? '');
  const aggregateAttrNames: string[] | null = hasAggregate ? getAliasNames(fetchXml ?? '') : null;

  attributesFieldNames.forEach((name, index) => {
    let displayName: string = displayNameCollection[name].DisplayName;

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

  linkEntityNames.forEach(async (linkEntityName, i) => {
    const linkAttributeNames: string[] = linkEntityAttributes[i].map(attr => attr.name);
    const mataData: EntityMetadata = await getEntityMetadata(linkEntityName, linkAttributeNames);

    linkEntityAttributes[i].forEach((attr, index) => {
      let fieldName = attr.attributeAlias;

      if (!fieldName && attr.linkEntityAlias) {
        fieldName = `${attr.linkEntityAlias}.${attr.name}`;
      }
      else if (!fieldName) {
        fieldName = `${linkEntityName}${i + 1}.${attr.name}`;
      }

      const columnName: string = attr.attributeAlias ||
       mataData.Attributes.get(attr.name).DisplayName;

      columns.push({
        name: columnName,
        fieldName,
        key: `col-el-${index}`,
        minWidth: 10,
        isResizable: true,
        isMultiline: false,
      });
    });
  });

  return columns;
};

export const openRecord = (entityName: string, entityId: string): void => {
  _context.navigation.openForm({
    entityName,
    entityId,
  });
};

export const openLookupForm = (
  entity: Entity,
  fieldName: string): void => {
  const entityName: string = entity[`_${fieldName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`];
  const entityId: string = entity[`_${fieldName}_value`];
  openRecord(entityName, entityId);
};

export const openLinkEntityRecord = (
  entity: Entity,
  fieldName: string): void => {
  const entityName: string = entity[`${fieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`];
  openRecord(entityName, entity[fieldName]);
};

export const openPrimaryEntityForm = (
  entity: Entity,
  entityName: string): void => {
  openRecord(entityName, entity[`${entityName}id`]);
};
