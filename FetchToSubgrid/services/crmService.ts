import * as React from 'react';
import { ColumnActionsMode, IColumn } from '@fluentui/react';
import { IFetchSubgridProps } from '../components/FetchSubgrid';
import { IInputs } from '../generated/ManifestTypes';
import {
  Dictionary,
  Entity,
  EntityAttribute,
  EntityMetadata,
  RetriveRecords } from '../utilities/types';
import {
  getEntityName,
  getLinkEntitiesNames,
  getAttributesFieldNames,
  isAggregate,
  getAliasNames } from '../utilities/utilities';

let _context: ComponentFramework.Context<IInputs>;

export const setContext = (context: ComponentFramework.Context<IInputs>) => {
  _context = context;
};

export const getProps = () => {
  const fetchXml = _context.parameters.fetchXmlProperty.raw;
  let pageSize = _context.parameters.defaultPageSize.raw || 1;
  if (pageSize && pageSize <= 0) {
    pageSize = 1;
  }

  try {
    const userParameters = JSON.parse(fetchXml ?? '');
    const props: IFetchSubgridProps = {
      fetchXml: userParameters.FetchXml || _context.parameters.defaultFetchXmlProperty.raw,
      defaultPageSize: Number(userParameters.PageSize) >= 250
        ? 250
        : Number(userParameters.PageSize) || pageSize,
      newButtonVisibility: userParameters.NewButtonVisibility ??
       _context.parameters.newButtonVisibility.raw === '0',
      deleteButtonVisibility: userParameters.DeleteButtonVisiblity ??
       _context.parameters.deleteButtonVisibility.raw === '0',
      userParameters,
    };
    return props;
  }
  catch {
    const props: IFetchSubgridProps = {
      fetchXml: fetchXml ?? _context.parameters.defaultFetchXmlProperty.raw,
      defaultPageSize: Number(pageSize) >= 250 ? 250 : Number(pageSize),
      newButtonVisibility: _context.parameters.newButtonVisibility.raw === '0',
      deleteButtonVisibility: _context.parameters.deleteButtonVisibility.raw === '0',
      userParameters: {},
    };
    return props;
  }
};

export const getEntityDisplayName = async (entityName: string): Promise<string> => {
  const entityMetadata = await _context.utils.getEntityMetadata(entityName);
  return entityMetadata._displayName;
};

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

  if (format === '3') {
    return _context.formatting.formatLanguage(fieldValue);
  }
  else if (format === '2') {
    return timeZoneDefinitions.value.find((e: any) =>
      e.timezonecode === Number(fieldValue)).userinterfacename;
  }
  else if (format === '0') {
    return fieldValue;
  }
  else if (fieldValue) {
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

  fetch?.removeAttribute('count');
  fetch?.removeAttribute('page');

  const fetchWithoutCount: string = new XMLSerializer().serializeToString(xmlDoc);
  const entityName: string = getEntityName(fetchWithoutCount);

  const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchWithoutCount ?? '')}`;
  const records: RetriveRecords =
  await _context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);

  let allRecordsCount = records.entities.length;
  let nextPage = 2;

  while (allRecordsCount >= 5000) {
    fetch.setAttribute('page', `${nextPage}`);
    const fetchNextPage: string = new XMLSerializer().serializeToString(xmlDoc);
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchNextPage ?? '')}`;
    const nextRecords: RetriveRecords =
    await _context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);

    allRecordsCount += nextRecords.entities.length;
    nextPage++;
    if (nextRecords.entities.length !== 5000) {
      return allRecordsCount;
    }
  }

  return allRecordsCount;
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
  const attributesFieldNames: string[] = getAttributesFieldNames(fetchXml ?? '');
  const entityName: string = getEntityName(fetchXml ?? '');

  const entityMetadata: EntityMetadata = await getEntityMetadata(entityName, attributesFieldNames);
  const displayNameCollection: Dictionary<EntityMetadata> =
   entityMetadata.Attributes._collection;

  const columns: IColumn[] = [];

  const linkEntityAttFieldNames: Dictionary<EntityAttribute[]> =
   getLinkEntitiesNames(fetchXml ?? '');
  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);

  const linkEntityAttributes: EntityAttribute[][] = Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => {
    const attributeNames: string[] = linkEntityAttributes[i].map(attr => attr.name);
    return getEntityMetadata(linkEntityNames, attributeNames);
  });
  const linkentityMetadata: EntityMetadata[] = await Promise.all(promises);

  const hasAggregate: boolean = isAggregate(fetchXml ?? '');
  const aggregateAttrNames: string[] | null = hasAggregate ? getAliasNames(fetchXml ?? '') : null;

  attributesFieldNames.forEach((name, index) => {
    let displayName = displayNameCollection[name].DisplayName;

    if (aggregateAttrNames?.length) {
      displayName = aggregateAttrNames[index];
      name = aggregateAttrNames[index];
    }

    columns.push({
      className: 'entity',
      name: displayName,
      fieldName: name,
      key: `col-${index}`,
      minWidth: 10,
      isResizable: true,
      isMultiline: false,
      columnActionsMode: ColumnActionsMode.hasDropdown,
    });
  });

  linkEntityNames.forEach((linkEntityName, i) => {
    linkEntityAttributes[i].forEach((attr, index) => {
      let fieldName = attr.attributeAlias;

      if (!fieldName && attr.linkEntityAlias) {
        fieldName = `${attr.linkEntityAlias}.${attr.name}`;
      }
      else if (!fieldName) {
        fieldName = `${linkEntityName}${i + 1}.${attr.name}`;
      }

      const columnName: string = attr.attributeAlias ||
      linkentityMetadata[i].Attributes._collection[attr.name].DisplayName;

      columns.push({
        className: 'linkEntity',
        ariaLabel: attr.name,
        name: columnName,
        fieldName,
        key: `col-el-${index}`,
        minWidth: 10,
        isResizable: true,
        isMultiline: false,
        columnActionsMode: ColumnActionsMode.hasDropdown,
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

const deleteSelectedRecords = async (recordIds: string[], entityName: string): Promise<void> => {
  try {
    for (const id of recordIds) {
      await _context.webAPI.deleteRecord(entityName, id);
    }
  }
  catch (e) {
    console.log(e);
  }
};

export const openRecordDeleteDialog = async (
  selectedRecordIds: string[],
  entityName: string,
  setDialogAccepted: React.Dispatch<React.SetStateAction<boolean>>): Promise<void> => {

  const entityMetadata = await _context.utils.getEntityMetadata(entityName);
  const confirmStrings = { text: `Do you want to delete this ${entityMetadata._displayName}?
   You can't undo this action.`, title: 'Confirm Deletion' };
  const confirmOptions = { height: 200, width: 450 };

  const deleteDialogStatus =
   await _context.navigation.openConfirmDialog(confirmStrings, confirmOptions);

  if (deleteDialogStatus.confirmed) {
    setDialogAccepted(true);
    await deleteSelectedRecords(selectedRecordIds, entityName);
    setDialogAccepted(false);
  }
};
