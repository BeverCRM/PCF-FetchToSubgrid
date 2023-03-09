import { IInputs } from '../generated/ManifestTypes';
import { WholeNumberType } from '../utilities/enums';
import { changeAliasNames,
  checkFetchXmlFormat,
  getEntityNameFromFetchXml,
} from '../utilities/fetchXmlUtils';
import {
  Entity,
  EntityMetadata,
  IAppWrapperProps,
  IDataverseService,
  JsonProps,
  RetriveRecords,
} from '../utilities/types';

export class DataverseService implements IDataverseService {
  private _context: ComponentFramework.Context<IInputs>;

  constructor(context: ComponentFramework.Context<IInputs>) {
    this._context = context;
  }

  private getPageSize(jsonObj?: JsonProps) {
    const pageSizse = Number(jsonObj?.pageSize);
    if (pageSizse) {
      if (pageSizse < 1) return 1;
      if (pageSizse > 250) return 250;
      return pageSizse;
    }

    const defaultPageSizse = this._context.parameters.defaultPageSize.raw ?? 0;
    if (defaultPageSizse < 1) return 1;
    if (defaultPageSizse > 250) return 250;
    return defaultPageSizse;
  }

  private async deleteSelectedRecords(recordIds: string[], entityName: string): Promise<void> {
    try {
      for (const id of recordIds) {
        await this._context.webAPI.deleteRecord(entityName, id);
      }
    }
    catch (e) {
      console.log(e);
    }
  }

  public getProps(): IAppWrapperProps {
    const fetchXmlorJson = this._context.parameters.fetchXmlProperty.raw;
    let pageSize = this._context.parameters.defaultPageSize.raw || 1;
    if (pageSize <= 0) pageSize = 1;

    try {
      const fieldValueJson: JsonProps = JSON.parse(fetchXmlorJson ?? '') as JsonProps;

      const allowedProps = [
        'newButtonVisibility',
        'deleteButtonVisibility',
        'pageSize',
        'fetchXml',
      ];

      const isJsonValid = Object.keys(fieldValueJson).every(prop => allowedProps.includes(prop));

      const props: IAppWrapperProps = {
        _service: this,
        fetchXml: fieldValueJson.fetchXml || this._context.parameters.defaultFetchXmlProperty.raw,
        defaultPageSize: this.getPageSize(fieldValueJson),
        allocatedWidth: this.getAllocatedWidth(),
        isJsonValid,
        newButtonVisibility: fieldValueJson.newButtonVisibility ??
          this._context.parameters.newButtonVisibility.raw === '1',
        deleteButtonVisibility: fieldValueJson.deleteButtonVisiblity ??
          this._context.parameters.deleteButtonVisibility.raw === '1',
        fetchXmlorJson,
      };

      return props;
    }
    catch {
      const isJsonValid = !!checkFetchXmlFormat(fetchXmlorJson);

      const props: IAppWrapperProps = {
        _service: this,
        fetchXml: fetchXmlorJson ?? this._context.parameters.defaultFetchXmlProperty.raw,
        defaultPageSize: this.getPageSize(),
        isJsonValid,
        allocatedWidth: this.getAllocatedWidth(),
        newButtonVisibility: this._context.parameters.newButtonVisibility.raw === '1',
        deleteButtonVisibility: this._context.parameters.deleteButtonVisibility.raw === '1',
        fetchXmlorJson,
      };
      return props;
    }
  }

  public async getEntityDisplayName(entityName: string): Promise<string> {
    const entityMetadata = await this._context.utils.getEntityMetadata(entityName);
    return entityMetadata._displayName;
  }

  public async getTimeZoneDefinitions(): Promise<void> {
    // @ts-ignore
    const contextPage = this._context.page;

    const request = await fetch(`${contextPage.getClientUrl()}/api/data/v9.0/timezonedefinitions`);
    const results = await request.json();

    return results;
  }

  public getWholeNumberFieldName(
    format: string,
    entity: Entity,
    fieldName: string,
    timeZoneDefinitions: any,
  ): string {
    let fieldValue: number = entity[fieldName];

    if (!fieldValue) return '';

    if (format === WholeNumberType.Language) {
      return this._context.formatting.formatLanguage(fieldValue);
    }

    if (format === WholeNumberType.TimeZone) {
      for (const tz of timeZoneDefinitions.value) {
        if (tz.timezonecode === Number(fieldValue)) return tz.userinterfacename;
      }
    }

    if (format === WholeNumberType.Duration) {
      let unit: string;

      if (fieldValue < 60) {
        unit = 'minute';
      }
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
  }

  public async getRecordsCount(fetchXml: string): Promise<number> {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
    const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

    fetch?.removeAttribute('count');
    fetch?.removeAttribute('page');

    const fetchWithoutCount: string = new XMLSerializer().serializeToString(xmlDoc);
    const changedAliasNames: string = changeAliasNames(fetchWithoutCount);
    const entityName: string = getEntityNameFromFetchXml(changedAliasNames);

    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(changedAliasNames ?? '')}`;
    const records: RetriveRecords = await this._context.webAPI.retrieveMultipleRecords(
      `${entityName}`,
      encodeFetchXml);

    let allRecordsCount: number = records.entities.length;
    let nextPage = 2;

    while (allRecordsCount >= 5000) {
      fetch.setAttribute('page', `${nextPage}`);
      const fetchNextPage: string = new XMLSerializer().serializeToString(xmlDoc);
      const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchNextPage ?? '')}`;
      const nextRecords: RetriveRecords =
        await this._context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);

      allRecordsCount += nextRecords.entities.length;
      nextPage++;
      if (nextRecords.entities.length !== 5000) {
        return allRecordsCount;
      }
    }

    return allRecordsCount;
  }

  public async getEntityMetadata(
    entityName: string,
    attributesFieldNames: string[],
  ): Promise<EntityMetadata> {
    const entityMetadata: EntityMetadata = await this._context.utils.getEntityMetadata(
      entityName,
      [...attributesFieldNames]);

    return entityMetadata;
  }

  public async getCurrentPageRecords(fetchXml: string | null): Promise<RetriveRecords> {
    const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
    return await this._context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
  }

  public openRecord(entityName: string, entityId: string): void {
    this._context.navigation.openForm({
      entityName,
      entityId,
    });
  }

  public openLookupForm(entity: Entity, fieldName: string): void {
    const entityName: string = entity[
      `_${fieldName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`];

    const entityId: string = entity[`_${fieldName}_value`];
    this.openRecord(entityName, entityId);
  }

  public openLinkEntityRecord(entity: Entity, fieldName: string): void {
    const entityName: string = entity[`${fieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`];
    this.openRecord(entityName, entity[fieldName]);
  }

  public openPrimaryEntityForm(entity: Entity, entityName: string): void {
    this.openRecord(entityName, entity[`${entityName}id`]);
  }

  public async openRecordDeleteDialog(
    selectedRecordIds: string[],
    entityName: string,
    setDialogAccepted: (value: any) => void,
  ): Promise<void> {
    const entityMetadata = await this._context.utils.getEntityMetadata(entityName);

    const confirmOptions = { height: 200, width: 450 };
    const confirmStrings = {
      text: `Do you want to delete this ${entityMetadata._displayName}?
      You can't undo this action.`, title: 'Confirm Deletion',
    };

    const deleteDialogStatus = await this._context.navigation.openConfirmDialog(
      confirmStrings,
      confirmOptions);

    if (deleteDialogStatus.confirmed) {
      setDialogAccepted(true);
      await this.deleteSelectedRecords(selectedRecordIds, entityName);
      setDialogAccepted(false);
    }
  }

  public async showNotificationPopup(error: any): Promise<void> {
    await this._context.navigation.openErrorDialog({
      message: error.message,
      details: error.stack,
    });
  }

  public getAllocatedWidth(): number {
    return this._context.mode.allocatedWidth;
  }
}
