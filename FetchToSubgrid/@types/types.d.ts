import { IColumn, ISelection } from '@fluentui/react';
import * as React from 'react';

export type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;

export type RetriveRecords = ComponentFramework.WebApi.RetrieveMultipleResponse;

export type Entity = ComponentFramework.WebApi.Entity;

export type Dictionary<T> = { [key: string]: T };

export interface IService<T> {
  _service: T;
}

export interface IDataverseService {
  getProps(): IAppWrapperProps;
  getEntityDisplayName(entityName: string): Promise<string>;
  getTimeZoneDefinitions(): Promise<void>;
  getWholeNumberFieldName(
    format: string,
    entity: Entity,
    fieldName: string,
    timeZoneDefinitions: any): string;
  getRecordsCount(fetchXml: string): Promise<number>;
  getEntityMetadata(entityName: string, attributesFieldNames: string[]): Promise<EntityMetadata>;
  getCurrentPageRecords(fetchXml: string | null): Promise<RetriveRecords>;
  openRecord(entityName: string, entityId: string): void;
  openLookupForm(entity: Entity, fieldName: string): void;
  openLinkEntityRecord(entity: Entity, fieldName: string): void;
  openPrimaryEntityForm(entity: Entity, entityName: string): void;
  showNotificationPopup(error: Error | undefined): Promise<void>;
  openRecordDeleteDialog(
    selectedRecordIds: string[],
    entityName: string,
    setDialogAccepted: (value: any) => void): Promise<void>;
  getAllocatedWidth(): number;
}

export interface EntityAttribute {
  linkEntityAlias?: string;
  name: string;
  attributeAlias: string;
}

export interface IItemProps {
  timeZoneDefinitions: any;
  item: Entity;
  isLinkEntity: boolean;
  entityMetadata: EntityMetadata;
  attributeType: number;
  fieldName: string;
  entity: Entity;
  pagingFetchData: string;
  index: number;
  hasAliasValue: boolean;
}

export interface IAppWrapperProps extends IService<IDataverseService> {
  fetchXml: string | null;
  pageSize: number;
  allocatedWidth: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
  error?: Error;
}

export interface IFetchToSubgridProps extends IAppWrapperProps {
  setIsLoading: (isLoading: boolean) => void;
  setError: (error?: Error | undefined) => void;
}

export interface IListProps extends IService<IDataverseService> {
  entityName: string;
  fetchXml: string | null;
  pageSize: number;
  inputsHashCode: number;
  currentPage: number;
  recordIds: React.MutableRefObject<string[]>;
  columns: IColumn[];
  items: Entity[];
  firstItemIndex: React.MutableRefObject<number>;
  lastItemIndex: React.MutableRefObject<number>;
  selectedItemsCount: React.MutableRefObject<number>;
  totalRecordsCount: number;
  nextButtonDisabled: React.MutableRefObject<boolean>;
  setItems: React.Dispatch<React.SetStateAction<ComponentFramework.WebApi.Entity[]>>;
  setColumns: React.Dispatch<React.SetStateAction<IColumn[]>>;
  setCurrentPage: React.Dispatch<React.SetStateAction<number>>;
  selection: ISelection;
}

export interface ICommandBarProps extends IService<IDataverseService> {
  isButtonActive: boolean;
  entityName: string;
  selectedRecordIds: string[];
  displayName: string;
  newButtonVisibility: boolean;
  deleteButtonVisibility: boolean | string;
  setDialogAccepted: React.Dispatch<React.SetStateAction<boolean>>;
}

export interface IFooterProps {
  firstItemIndex: number;
  lastItemIndex: number;
  selectedItemsCount: number;
  totalRecordsCount: number;
  currentPage: number;
  nextButtonDisable: boolean;
  movePreviousIsDisabled: boolean;
  setCurrentPage: (page: number) => void;
}

export interface IInfoMessageProps {
  error?: Error;
  dataverseService: IDataverseService;
}

export interface ILinkableItemProps extends IService<IDataverseService> {
  item: Entity
}
