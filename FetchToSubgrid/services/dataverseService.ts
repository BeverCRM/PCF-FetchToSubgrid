import { IInputs } from '../generated/ManifestTypes';
import { WholeNumberType } from '../@types/enums';
import { IAppWrapperProps } from '../components/AppWrapper';
import { changeAliasNames, getEntityNameFromFetchXml } from '../utilities/fetchXmlUtils';
import {
  Entity,
  EntityMetadata,
  RetriveRecords,
  DialogResponse,
} from '../@types/types';

export interface IDataverseService {
  getProps(): IAppWrapperProps;
  getEntityDisplayName(entityName: string): Promise<string>;
  getTimeZoneDefinitions(): Promise<Object>;
  getWholeNumberFieldName(
    format: string,
    entity: Entity,
    fieldName: string,
    timeZoneDefinitions: any): string;
  getRecordsCount(fetchXml: string): Promise<number>;
  getEntityMetadata(entityName: string, attributesFieldNames: string[]): Promise<EntityMetadata>;
  getCurrentPageRecords(fetchXml: string | null): Promise<RetriveRecords>;
  openRecordForm(entityName: string, entityId: string): void;
  openLookupForm(entity: Entity, fieldName: string): void;
  openLinkEntityRecordForm(entity: Entity, fieldName: string): void;
  openPrimaryEntityForm(entity: Entity, entityName: string): void;
  openErrorDialog(error: Error): Promise<void>;
  openRecordDeleteDialog(entityName: string): Promise<DialogResponse>;
  getAllocatedWidth(): number;
  deleteSelectedRecords(
    selectedRecordIds: string[],
    entityName: string,
  ): Promise<void>;
}

export class DataverseService implements IDataverseService {
  private _context: ComponentFramework.Context<IInputs>;

  constructor(context: ComponentFramework.Context<IInputs>) {
    this._context = context;
  }

  public getProps(): IAppWrapperProps {
    const fetchXmlOrJson: string | null = this._context.parameters.fetchXmlProperty.raw;

    let defaultPageSize: number = this._context.parameters.defaultPageSize.raw || 1;
    if (defaultPageSize < 0) defaultPageSize = 1;

    const props: IAppWrapperProps = {
      _service: this,
      fetchXmlOrJson,
      allocatedWidth: this.getAllocatedWidth(),
      default: {
        fetchXml: this._context.parameters.defaultFetchXmlProperty.raw,
        pageSize: defaultPageSize,
        newButtonVisibility: this._context.parameters.newButtonVisibility.raw === '1',
        deleteButtonVisibility: this._context.parameters.deleteButtonVisibility.raw === '1',
      },

    };
    return props;
  }

  public async getEntityDisplayName(entityName: string): Promise<string> {
    const entityMetadata: EntityMetadata = await this._context.utils.getEntityMetadata(entityName);
    return entityMetadata._displayName;
  }

  public async getTimeZoneDefinitions(): Promise<Object> {
    // @ts-ignore
    const contextPage = this._context.page;

    const response: Response = await fetch(
      `${contextPage.getClientUrl()}/api/data/v9.0/timezonedefinitions`);

    const results = await response.json();
    return results;
  }

  public getWholeNumberFieldName(
    format: string,
    entity: Entity,
    fieldName: string,
    timeZoneDefinitions: any,
  ): string {
    let fieldValue: number = entity[fieldName];
    let unit: string;

    if (!fieldValue) return '';

    switch (format) {
      case WholeNumberType.Language:
        return this._context.formatting.formatLanguage(fieldValue);

      case WholeNumberType.TimeZone:
        for (const tz of timeZoneDefinitions.value) {
          if (tz.timezonecode === Number(fieldValue)) return tz.userinterfacename;
        }
        return '';

      case WholeNumberType.Duration:
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

      case WholeNumberType.Number:
        return `${fieldValue}`;

      default:
        return '';
    }
  }

  public async getRecordsCount(fetchXml: string): Promise<number> {
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

  public async getCurrentPageRecords(fetchXml: string | null): Promise<RetriveRecords> {
    const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
    return await this._context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
  }

  public openRecordForm(entityName: string, entityId: string): void {
    this._context.navigation.openForm({
      entityName,
      entityId,
    });
  }

  public openLookupForm(entity: Entity, fieldName: string): void {
    const entityName: string = entity[
      `_${fieldName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`];

    const entityId: string = entity[`_${fieldName}_value`];
    this.openRecordForm(entityName, entityId);
  }

  public openLinkEntityRecordForm(entity: Entity, fieldName: string): void {
    const entityName: string = entity[`${fieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`];
    this.openRecordForm(entityName, entity[fieldName]);
  }

  public openPrimaryEntityForm(entity: Entity, entityName: string): void {
    this.openRecordForm(entityName, entity[`${entityName}id`]);
  }

  public async openRecordDeleteDialog(entityName: string): Promise<DialogResponse> {
    const entityMetadata: EntityMetadata = await this._context.utils.getEntityMetadata(entityName);

    const confirmOptions = { height: 200, width: 450 };
    const confirmStrings = {
      text: `Do you want to delete this ${entityMetadata._displayName}?
      You can't undo this action.`, title: 'Confirm Deletion',
    };

    const deleteDialogStatus: DialogResponse = await this._context.navigation.openConfirmDialog(
      confirmStrings, confirmOptions);

    return deleteDialogStatus;
  }

  public async openErrorDialog(error: Error): Promise<void> {
    await this._context.navigation.openErrorDialog({
      message: error?.message,
      details: error?.stack,
    });
  }

  public getAllocatedWidth(): number {
    return this._context.mode.allocatedWidth;
  }

  public async deleteSelectedRecords(recordIds: string[], entityName: string): Promise<void> {
    try {
      for (const id of recordIds) {
        await this._context.webAPI.deleteRecord(entityName, id);
      }
    }
    catch (e) {
      console.log(e);
    }
  }
}
