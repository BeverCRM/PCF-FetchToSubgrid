import * as React from 'react';
import { IColumn } from '@fluentui/react';
import { genereateItems, getColumns, getItems, IItemProps } from './d365Utils';
import { AttributeType } from '../@types/enums';
import { getFetchXmlParserError, getOrderInFetchXml } from './fetchXmlUtils';
import { LinkableItem } from '../components/LinkableItems';
import { IDataverseService } from '../services/dataverseService';
import { IAppWrapperProps } from '../components/AppWrapper';
import { IFetchToSubgridProps } from '../components/FetchToSubgrid';
import {
  Entity,
  EntityAttribute,
  IItemsData,
  IRecordsData,
  OrderInFetchXml } from '../@types/types';

interface IJsonProps {
  newButtonVisibility: boolean;
  deleteButtonVisibility: boolean;
  pageSize: number;
  fetchXml: string;
}

type JsonAllowedProps = Array<keyof IJsonProps>;

export const checkIfAttributeIsEntityReference = (attributeType: AttributeType): boolean => {
  const attributetypes: AttributeType[] = [
    AttributeType.Lookup,
    AttributeType.Owner,
    AttributeType.Customer,
  ];

  return attributetypes.includes(attributeType);
};

export const checkIfAttributeRequiresFormattedValue = (attributeType: AttributeType) => {
  const attributeTypes: AttributeType[] = [
    AttributeType.Money,
    AttributeType.PickList,
    AttributeType.DateTime,
    AttributeType.MultiselectPickList,
    AttributeType.TwoOptions,
  ];

  return attributeTypes.includes(attributeType);
};

export const sortColumns = (
  allColumns: IColumn[],
  fieldName?: string,
  descending?: boolean): IColumn[] => {
  const filteredColumns = allColumns.map(col => {
    const isFieldMatch = fieldName !== undefined && col.fieldName === fieldName;

    col.isSorted = isFieldMatch;

    if (isFieldMatch) {
      col.isSortedDescending = descending !== undefined ? descending : !col.isSortedDescending;
    }

    return col;
  });

  return filteredColumns;
};

export const getSortedColumns = async (
  fetchXml: string | null,
  allocatedWidth: number,
  dataverseService: IDataverseService): Promise<IColumn[]> => {
  const columns: IColumn[] = await getColumns(fetchXml, allocatedWidth, dataverseService);
  const order: OrderInFetchXml | null = getOrderInFetchXml(fetchXml ?? '');

  if (!order) return columns;

  const filteredColumns: IColumn[] = sortColumns(
    columns,
    Object.keys(order)[0],
    Object.values(order)[0]);

  return filteredColumns;
};

const defaultProps: IJsonProps = {
  newButtonVisibility: false,
  deleteButtonVisibility: false,
  pageSize: 0,
  fetchXml: '',
};

export const isJsonValid = (jsonObj: Object): boolean => {
  const allowedProps: JsonAllowedProps = Object.keys(defaultProps) as JsonAllowedProps;
  return Object.keys(jsonObj).every(prop => allowedProps.includes(prop as keyof IJsonProps));
};

export const createLinkableItems = (
  records: Entity[],
  dataverseService: IDataverseService): Entity[] => {
  records.forEach(record => {
    Object.keys(record).forEach(key => {
      if (key !== 'id') {
        const value: any = record[key];

        record[key] = value.isLinkable ? React.createElement(LinkableItem, {
          _service: dataverseService,
          item: value,
        })
          : value.displayName;
      }
    });
  });

  return records;
};

export const getPageSize = (value?: IJsonProps | number | null): number => {
  let pageSize = 1;

  if (typeof value === 'number') pageSize = value;
  else if (value) pageSize = Number(value?.pageSize);

  if (isNaN(pageSize) || pageSize < 1) return 1;
  if (pageSize > 250) return 250;
  return pageSize;
};

export const parseRawInput = (appWrapperProps: IAppWrapperProps) => {
  const { fetchXmlOrJson } = appWrapperProps;

  const props: IFetchToSubgridProps = {
    error: undefined,
    _service: appWrapperProps._service,
    allocatedWidth: appWrapperProps.allocatedWidth,
    fetchXml: appWrapperProps.default.fetchXml,
    pageSize: getPageSize(appWrapperProps.default.pageSize),
    newButtonVisibility: appWrapperProps.default.newButtonVisibility,
    deleteButtonVisibility: appWrapperProps.default.deleteButtonVisibility,
    setIsLoading: () => {},
    setError: () => {},
  };

  try {
    const fieldValueJson: IJsonProps = JSON.parse(fetchXmlOrJson ?? '') as IJsonProps;
    if (!isJsonValid(fieldValueJson)) props.error = new Error('JSON is not valid');

    if (fieldValueJson.fetchXml) props.fetchXml = fieldValueJson.fetchXml;
    if (fieldValueJson.pageSize) props.pageSize = fieldValueJson.pageSize;

    if (fieldValueJson.newButtonVisibility !== undefined) {
      props.newButtonVisibility = fieldValueJson.newButtonVisibility;
    }

    if (fieldValueJson.deleteButtonVisibility !== undefined) {
      props.deleteButtonVisibility = fieldValueJson.deleteButtonVisibility;
    }
  }
  catch {
    const fieldValueFetchXml: string | null = fetchXmlOrJson || appWrapperProps.default.fetchXml;
    const fetchXmlParserError: string | null = getFetchXmlParserError(fieldValueFetchXml);

    if (fetchXmlParserError) props.error = new Error(fetchXmlParserError);
    if (fieldValueFetchXml) props.fetchXml = fieldValueFetchXml;
  }

  return props;
};

export const hashCode = (str: string) => {
  if (str.length === 0) return 0;

  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const charCode = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + charCode;
    hash |= 0; // Convert to 32bit integer
  }
  return hash;
};

export const getFormattingFieldValue = (fieldValue: number): string => {
  let unit: string;

  if (fieldValue < 60) {
    return 'minute';
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
};

export const getLinkableItems = async (
  data: IItemsData,
  dataverseService: IDataverseService): Promise<Entity[]> => {
  const records: Entity[] = await getItems(
    data.fetchXml,
    data.pageSize,
    data.currentPage,
    dataverseService);

  const linkableItems = createLinkableItems(records, dataverseService);

  return linkableItems;
};

export const getAttributeAliasName = (
  attribute: EntityAttribute,
  index: number,
  linkEntityName: string) => {
  if (!attribute.attributeAlias && attribute.linkEntityAlias) {
    return `${attribute.linkEntityAlias}.${attribute.name}`;
  }
  else if (!attribute.attributeAlias) {
    return `${linkEntityName}${index + 1}.${attribute.name}`;
  }

  return attribute.attributeAlias;
};

export const genereateItemsForEntity = (
  recordsData: IRecordsData, item: Entity, entity: Entity, dataverseService: IDataverseService) => {

  recordsData.attributesFieldNames.forEach((fieldName: any, index: number) => {
    const hasAliasValue = !!recordsData.entityAliases[index];
    const attributeType: number = recordsData.entityMetadata.Attributes.get(
      fieldName).AttributeType;

    if (recordsData.entityAliases[index]) {
      fieldName = recordsData.entityAliases[index];
    }

    const attributes: IItemProps = {
      timeZoneDefinitions: recordsData.timeZoneDefinitions,
      item,
      isLinkEntity: false,
      entityMetadata: recordsData.entityMetadata,
      attributeType,
      fieldName,
      entity,
      pagingFetchData: recordsData.pagingFetchData,
      index,
      hasAliasValue,
    };

    genereateItems(attributes, dataverseService);
  });
};

export const genereateItemsForLinkEntity = (
  recordsData: IRecordsData, item: Entity, entity: Entity, dataverseService: IDataverseService) => {
  recordsData.linkEntityNames.forEach((linkEntityName: string, i: number) => {
    recordsData.linkEntityAttributes[i].forEach((attr: any, index: number) => {
      const hasAliasValue = !!attr.linkEntityAlias;
      const attributeType: number = recordsData.linkentityMetadata[i].Attributes.get(
        attr.name).AttributeType;
      const fieldName = getAttributeAliasName(attr, i, linkEntityName);

      const attributes: IItemProps = {
        timeZoneDefinitions: recordsData.timeZoneDefinitions,
        item,
        isLinkEntity: true,
        entityMetadata: recordsData.linkentityMetadata[i],
        attributeType,
        fieldName,
        entity,
        pagingFetchData: recordsData.pagingFetchData,
        index,
        hasAliasValue,
      };

      genereateItems(attributes, dataverseService);
    });
  });
};
