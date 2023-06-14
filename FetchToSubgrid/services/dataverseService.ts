/* eslint-disable prefer-destructuring */
import { IInputs } from '../generated/ManifestTypes';
import { WholeNumberType } from '../@types/enums';
import { IAppWrapperProps } from '../components/AppWrapper';
import { getFormattingFieldValue } from '../utilities/utils';
import { Entity, EntityMetadata, RetrieveRecords, DialogResponse } from '../@types/types';
import {
  changeAliasNames,
  getEntityNameFromFetchXml,
  addAttributeIdInFetch } from '../utilities/fetchXmlUtils';

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
  getCurrentPageRecords(fetchXml: string | null): Promise<RetrieveRecords>;
  openRecordForm(entityName: string, entityId: string): void;
  openLookupForm(entity: Entity, fieldName: string): void;
  openLinkEntityRecordForm(entity: Entity, fieldName: string): void;
  openPrimaryEntityForm(entity: Entity, entityName: string): void;
  openErrorDialog(error: Error): Promise<void>;
  openRecordDeleteDialog(entityName: string): Promise<DialogResponse>;
  getRelationships(parentEntityName: string): any;
  openNewRecord(entityName: string): void;
  deleteSelectedRecords(
    selectedRecordIds: string[],
    entityName: string,
  ): Promise<void>;
}

type RelationshipEntity = {
  ReferencingEntity: string;
  ReferencingAttribute: string;
  MetadataId: string;
};

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
        fetchXml: this._context.parameters.defaultFetchXml.raw,
        pageSize: defaultPageSize,
        newButtonVisibility: this._context.parameters.newButtonVisibility.raw === '1',
        deleteButtonVisibility: this._context.parameters.deleteButtonVisibility.raw === '1',
      },

    };
    return props;
  }

  public async getEntityDisplayName(entityName: string): Promise<string> {
    const entityMetadata: EntityMetadata = await this._context.utils.getEntityMetadata(entityName);
    return entityMetadata?._displayName;
  }

  public async getTimeZoneDefinitions(): Promise<Object> {
    // @ts-ignore
    const contextPage = this._context.page;

    const response: Response = await fetch(
      `${contextPage.getClientUrl()}/api/data/v9.0/timezonedefinitions`);

    const results = await response.json();
    return results;
  }

  public async openNewRecord(entityName: string): Promise<void> {
    // @ts-ignore
    const contextPage = this._context.page;
    // @ts-ignore
    const lookupName: string = this._context.mode.contextInfo.entityRecordName;
    const parentEntityName: string = contextPage.entityTypeName;
    const entityId: string = contextPage.entityId;

    const relationshipEntities: RelationshipEntity[] = await this.getRelationships(
      parentEntityName);

    const relationship: RelationshipEntity | undefined = relationshipEntities.find(
      relationshipEntity => relationshipEntity.ReferencingEntity === entityName);

    if (relationship) {
      const lookup: { id: string, name: string, entityType: string } =
      { id: entityId, name: lookupName, entityType: parentEntityName };
      const formParameters: { [key: string]: string } =
       { [relationship.ReferencingAttribute]: JSON.stringify(lookup) };

      this._context.navigation.openForm({ entityName }, formParameters);
    }
    else {
      this._context.navigation.openForm({ entityName });
    }
  }

  public async getRelationships(parentEntityName: string): Promise<RelationshipEntity[]> {
    // @ts-ignore
    const contextPage = this._context.page;
    const entityDefinitions = `/api/data/v9.2/EntityDefinitions(LogicalName='${parentEntityName}')`;
    const relationships = `OneToManyRelationships($select=ReferencingEntity,ReferencingAttribute)`;
    const logicalName = '&$select=LogicalName';

    const response: Response = await fetch(
      `${contextPage.getClientUrl()}${entityDefinitions}?$expand=${relationships}${logicalName}`);

    const result = await response.json();
    return result.OneToManyRelationships;
  }

  public getWholeNumberFieldName(
    format: string,
    entity: Entity,
    fieldName: string,
    timeZoneDefinitions: any,
  ): string {
    const fieldValue: number = entity[fieldName];

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
        return getFormattingFieldValue(fieldValue);

      case WholeNumberType.Number:
        return `${fieldValue}`;

      default:
        return '';
    }
  }

  public async getRecordsCount(fetchXml: string): Promise<number> {
    const top: number = this.getTopInFetchXml(fetchXml);
    if (top) return top;

    let numberOfRecords = 0;
    let page = 0;
    let pagingCookie: string | null = null;

    const entityName: string = getEntityNameFromFetchXml(fetchXml);
    const changedAliasNames: string = changeAliasNames(fetchXml);

    const updateFetchXml = (xml: string, page: number): string => {
      const parser: DOMParser = new DOMParser();
      const xmlDoc: Document = parser.parseFromString(xml, 'text/xml');
      const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

      fetch?.removeAttribute('top');
      fetch?.removeAttribute('count');
      fetch?.removeAttribute('page');

      fetch?.setAttribute('page', `${page}`);

      const serializer = new XMLSerializer();
      return serializer.serializeToString(xmlDoc);
    };

    do {
      const updatedFetchXml: string = updateFetchXml(changedAliasNames, ++page);
      const data: any = await this._context.webAPI.retrieveMultipleRecords(
        entityName, `?fetchXml=${encodeURIComponent(updatedFetchXml)}`);

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

  public async getCurrentPageRecords(fetchXml: string | null): Promise<RetrieveRecords> {
    const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
    const updatedFetchXml: string = addAttributeIdInFetch(fetchXml ?? '', entityName);
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(updatedFetchXml ?? '')}`;

    return await this._context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
  }

  public openRecordForm(entityName: string, entityId: string): void {
    this._context.navigation.openForm({
      entityName,
      entityId,
    });
  }

  public openLookupForm(entity: Entity, fieldName: string): void {
    let entityName: string;
    let entityId: string;
    if (fieldName.startsWith('alias')) {
      entityName = entity[`${fieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`];
      entityId = entity[`${fieldName}`];
    }
    else {
      entityName = entity[`_${fieldName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`];
      entityId = entity[`_${fieldName}_value`];
    }
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
      text: `Do you want to delete this ${entityMetadata?._displayName}?
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

  private getAllocatedWidth(): number {
    return this._context.mode.allocatedWidth;
  }

  private getTopInFetchXml(fetchXml: string): number {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

    const top: string | null | undefined = xmlDoc.querySelector('fetch')?.getAttribute('top');

    if (top) {
      return Number(top);
    }
    return 0;
  }
}
