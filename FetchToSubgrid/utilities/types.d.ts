import { IColumn } from '@fluentui/react';
import * as React from 'react';

export type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;

export type RetriveRecords = ComponentFramework.WebApi.RetrieveMultipleResponse;

export type Entity = ComponentFramework.WebApi.Entity;

export type Dictionary<T> = { [key: string]: T };

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

export interface IAppWrapperProps {
  fetchXml: string | null;
  defaultPageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
}

export interface IFetchSubgridProps {
  fetchXml: string | null;
  defaultPageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
  setIsLoading: (isLoading: boolean) => void;
  setErrorMessage: (message?: string) => void;
  isVisible: boolean;
}

export interface IListProps {
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

export interface ICommandBarProps {
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

export interface ILinkableItemProps {
  item: Entity
}
