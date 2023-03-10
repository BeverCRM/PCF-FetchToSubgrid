import { IInputs } from '../generated/ManifestTypes';
import { WholeNumberType } from '../utilities/enums';
import {
  changeAliasNames,
  getEntityNameFromFetchXml,
  getFetchXmlParserError,
} from '../utilities/fetchXmlUtils';
import {
  Entity,
  EntityMetadata,
  IAppWrapperProps,
  IDataverseService,
  JsonProps,
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
    const fetchXmlOrJson = this._context.parameters.fetchXmlProperty.raw;
    let pageSize = this._context.parameters.defaultPageSize.raw || 1;
    if (pageSize <= 0) pageSize = 1;

    let error: Error | undefined = undefined;

    try {
      const allowedProps = [
        'newButtonVisibility',
        'deleteButtonVisibility',
        'pageSize',
        'fetchXml',
      ];

      const fieldValueJson = JSON.parse(fetchXmlOrJson ?? '') as JsonProps;
      const isJsonValid = Object.keys(fieldValueJson).every(prop => allowedProps.includes(prop));
      if (!isJsonValid) error = new Error('JSON is not valid');

      const props: IAppWrapperProps = {
        _service: this,
        fetchXml: fieldValueJson.fetchXml || this._context.parameters.defaultFetchXmlProperty.raw,
        pageSize: this.getPageSize(fieldValueJson),
        allocatedWidth: this.getAllocatedWidth(),
        error,
        newButtonVisibility: fieldValueJson.newButtonVisibility ??
          this._context.parameters.newButtonVisibility.raw === '1',
        deleteButtonVisibility: fieldValueJson.deleteButtonVisibility ??
          this._context.parameters.deleteButtonVisibility.raw === '1',
      };

      return props;
    }
    catch {
      const fetchXml = fetchXmlOrJson ?? this._context.parameters.defaultFetchXmlProperty.raw;
      const fetchXmlParserError: string | null = getFetchXmlParserError(fetchXml);

      if (fetchXmlParserError) error = new Error(fetchXmlParserError);

      const props: IAppWrapperProps = {
        _service: this,
        fetchXml,
        pageSize: this.getPageSize(),
        error,
        allocatedWidth: this.getAllocatedWidth(),
        newButtonVisibility: this._context.parameters.newButtonVisibility.raw === '1',
        deleteButtonVisibility: this._context.parameters.deleteButtonVisibility.raw === '1',
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

  public getRecordsCount = async (fetchXml: string) => {
    let pagingCookie = null;
    let numberOfRecords = 0;
    let page = 0;
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
    const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

    const entityName: string = getEntityNameFromFetchXml(fetchXml);
    fetch?.removeAttribute('count');
    const changedAliasNames: string = changeAliasNames(fetchXml);

    do {
      fetch?.removeAttribute('page');
      fetch.setAttribute('page', `${++page}`);
      // eslint-disable-next-line no-invalid-this
      const data: any = await this._context.webAPI.retrieveMultipleRecords(
        entityName, `?fetchXml=${encodeURIComponent(changedAliasNames)}`);
      numberOfRecords += data.entities.length;
      pagingCookie = data.fetchXmlPagingCookie;

    } while (pagingCookie);

    return numberOfRecords;
  }

  public async getEntityMetadata(
    entityName: string, attributesFieldNames: string[]): Promise<EntityMetadata> {
    const entityMetadata: EntityMetadata = await this._context.utils.getEntityMetadata(
      entityName,
      [...attributesFieldNames]);

    return entityMetadata;
  }

  public async getCurrentPageRecords(fetchXml: string | null):
   Promise<any> {
    try {
      const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
      const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
      return await this._context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
    }
    catch (error) {
      return error;
    }
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

  public async showNotificationPopup(error: Error | undefined): Promise<void> {
    await this._context.navigation.openErrorDialog({
      message: error?.message,
      details: error?.stack,
    });
  }

  public getAllocatedWidth(): number {
    return this._context.mode.allocatedWidth;
  }
}
