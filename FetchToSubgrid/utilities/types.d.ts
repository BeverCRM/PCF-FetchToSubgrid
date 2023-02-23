import { IColumn } from '@fluentui/react';
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
  openRecordDeleteDialog(
    selectedRecordIds: string[],
    entityName: string,
    setDialogAccepted: (value: any) => void): Promise<void>;
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
  fetchXml: string | null;
  index: number;
}

export interface IAppWrapperProps extends IService<IDataverseService> {
  fetchXml: string | null;
  defaultPageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
}

export interface IFetchSubgridProps extends IAppWrapperProps {
  setIsLoading: (isLoading: boolean) => void;
  setErrorMessage: (message?: string) => void;
  isVisible: boolean;
}

export interface IListProps extends IService<IDataverseService> {
  entityName: string;
  fetchXml: string | null;
  pageSize: number;
  currentPage: number;
  recordIds: React.MutableRefObject<string[]>;
  columns: IColumn[];
  items: Entity[];
  deleteBtnClassName: React.MutableRefObject<string>
  firstItemIndex: React.MutableRefObject<number>;
  lastItemIndex: React.MutableRefObject<number>;
  selectedItemsCount: React.MutableRefObject<number>;
  totalRecordsCount: React.MutableRefObject<number>;
  nextButtonDisabled: React.MutableRefObject<boolean>;
  setItems: React.Dispatch<React.SetStateAction<ComponentFramework.WebApi.Entity[]>>;
  setColumns: React.Dispatch<React.SetStateAction<IColumn[]>>;
  setCurrentPage: React.Dispatch<React.SetStateAction<number>>;
  setSelectedRecordIds: React.Dispatch<React.SetStateAction<string[]>>
}

export interface ICommandBarProps extends IService<IDataverseService> {
  className: string;
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
  selectedItems: number;
  totalRecordsCount: number;
  currentPage: number;
  nextButtonDisable: boolean;
  movePreviousIsDisabled: boolean;
  setCurrentPage: (page: number) => void;
}

export interface IInfoMessageProps {
  message?: string;
}

export interface ILinkableItemProps extends IService<IDataverseService> {
  item: Entity
}

export interface JsonProps {
  newButtonVisibility?: boolean;
  deleteButtonVisiblity?: boolean;
  pageSize?: number;
  fetchXml?: string;
}
