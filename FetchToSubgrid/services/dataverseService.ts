import { IInputs } from '../generated/ManifestTypes';
import { WholeNumberType } from '../utilities/enums';
import { getEntityNameFromFetchXml } from '../utilities/fetchXmlUtils';
import {
  Entity,
  EntityMetadata,
  IAppWrapperProps,
  RetriveRecords } from '../utilities/types';

let _context: ComponentFramework.Context<IInputs>;

export const setContext = (context: ComponentFramework.Context<IInputs>) => {
  _context = context;
};

const getPageSize = (jsonObj?: any) => {
  const pageSizse = Number(jsonObj?.PageSize);
  if (pageSizse) {
    if (pageSizse < 1) return 1;
    if (pageSizse > 250) return 250;
    return pageSizse;
  }

  const defaultPageSizse = _context.parameters.defaultPageSize.raw ?? 0;
  if (defaultPageSizse < 1) return 1;
  if (defaultPageSizse > 250) return 250;
  return defaultPageSizse;
};

export const getProps = (): IAppWrapperProps => {
  const fetchXml = _context.parameters.fetchXmlProperty.raw;
  let pageSize = _context.parameters.defaultPageSize.raw || 1;
  if (pageSize <= 0) pageSize = 1;

  try {
    const fieldValueJson = JSON.parse(fetchXml ?? '');

    const props: IAppWrapperProps = {
      fetchXml: fieldValueJson.FetchXml || _context.parameters.defaultFetchXmlProperty.raw,
      defaultPageSize: getPageSize(fieldValueJson),
      newButtonVisibility: fieldValueJson.NewButtonVisibility ??
        _context.parameters.newButtonVisibility.raw === '0',
      deleteButtonVisibility: fieldValueJson.DeleteButtonVisiblity ??
       _context.parameters.deleteButtonVisibility.raw,
    };
    return props;
  }
  catch {
    const props: IAppWrapperProps = {
      fetchXml: fetchXml ?? _context.parameters.defaultFetchXmlProperty.raw,
      defaultPageSize: getPageSize(),
      newButtonVisibility: _context.parameters.newButtonVisibility.raw === '0',
      deleteButtonVisibility: _context.parameters.deleteButtonVisibility.raw === '0',
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
  timeZoneDefinitions: any): string => {
  let fieldValue: number = entity[fieldName];

  if (!fieldValue) return '';

  if (format === WholeNumberType.Language) {
    return _context.formatting.formatLanguage(fieldValue);
  }

  if (format === WholeNumberType.TimeZone) {
    for (const tz of timeZoneDefinitions.value) {
      if (tz.timezonecode === Number(fieldValue)) return tz.userinterfacename;
    }
  }

  if (format === WholeNumberType.Duration) {
    let unit: string;
    if (fieldValue < 60) { unit = 'minute'; }
    else if (fieldValue < 1440) {
      fieldValue = Math.round(fieldValue / 60 * 100) / 100;
      unit = 'hour';
    }
    else {
      Math.round(fieldValue / 1440 * 100) / 100;
      unit = 'day';
    }
    return `${fieldValue} ${unit}${fieldValue === 1 ? '' : 's'}`;
  }

  if (format === WholeNumberType.Number) {
    return `${fieldValue}`;
  }

  return '';
};

export const getRecordsCount = async (fetchXml: string): Promise<number> => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

  fetch?.removeAttribute('count');
  fetch?.removeAttribute('page');

  const fetchWithoutCount: string = new XMLSerializer().serializeToString(xmlDoc);
  const entityName: string = getEntityNameFromFetchXml(fetchWithoutCount);

  const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchWithoutCount ?? '')}`;
  const records: RetriveRecords = await _context.webAPI.retrieveMultipleRecords(
    `${entityName}`,
    encodeFetchXml);

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

export const getCurrentPageRecords = async (fetchXml: string | null): Promise<RetriveRecords> => {
  const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
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
  setDialogAccepted: (value: any) => void,
): Promise<void> => {

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
