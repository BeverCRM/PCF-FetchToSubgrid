import * as React from 'react';
import { IColumn } from '@fluentui/react';
import { getColumns } from './d365Utils';
import { AttributeType } from '../@types/enums';
import { getFetchXmlParserError, getOrderInFetchXml } from './fetchXmlUtils';
import { LinkableItem } from '../components/LinkableItems';
import {
  Entity,
  IAppWrapperProps,
  IDataverseService,
  IFetchToSubgridProps,
  IJsonProps,
  JsonAllowedProps,
  OrderInFetchXml,
} from '../@types/types';

export const checkIfAttributeIsEntityReferance = (attributeType: AttributeType): boolean => {
  const attributetypes: AttributeType[] = [
    AttributeType.Lookup,
    AttributeType.Owner,
    AttributeType.Customer,
  ];

  return attributetypes.includes(attributeType);
};

export const needToGetFormattedValue = (attributeType: AttributeType) => {
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
  fieldName?: string,
  ariaLabel?: string,
  descending?: boolean,
  allColumns?: IColumn[]): IColumn[] => {
  const filteredColumns = allColumns?.map((col: IColumn) => {
    const fetchXmlOrderMathAttributeFieldName = fieldName === col.fieldName &&
      fieldName !== undefined;
    const fetchXmlOrderMathcLinkEntityAriaLabel = ariaLabel === col.ariaLabel &&
      ariaLabel !== undefined;

    col.isSorted = fetchXmlOrderMathAttributeFieldName || fetchXmlOrderMathcLinkEntityAriaLabel;

    if (fetchXmlOrderMathAttributeFieldName || fetchXmlOrderMathcLinkEntityAriaLabel) {
      col.isSortedDescending = descending !== undefined ? descending : !col.isSortedDescending;
    }

    return col;
  });

  return filteredColumns ?? [];
};

export const getSortedColumns = async (
  fetchXml: string | null,
  allocatedWidth: number,
  dataverseService: IDataverseService): Promise<IColumn[]> => {
  const columns: IColumn[] = await getColumns(fetchXml, allocatedWidth, dataverseService);
  const order: OrderInFetchXml | null = getOrderInFetchXml(fetchXml ?? '');

  if (!order) return columns;

  const filteredColumns: IColumn[] = sortColumns(
    Object.keys(order)[0],
    Object.keys(order)[0],
    Object.values(order)[0],
    columns);

  return filteredColumns;
};

export const calculateFilteredRecordsData = (
  totalRecordsCount: number,
  records: Entity[],
  pageSize: number,
  currentPage: number,
  nextButtonDisabled: React.MutableRefObject<boolean>,
  lastItemIndex: React.MutableRefObject<number>,
  firstItemIndex: React.MutableRefObject<number>): void => {
  nextButtonDisabled.current = Math.ceil(totalRecordsCount / pageSize) <= currentPage;

  lastItemIndex.current = (currentPage - 1) * pageSize + records.length;
  firstItemIndex.current = (currentPage - 1) * pageSize + 1;
};

class JsonProps implements IJsonProps {
  newButtonVisibility: boolean = false;
  deleteButtonVisibility: boolean = false;
  pageSize: number = 0;
  fetchXml: string = '';
}

export const isJsonValid = (jsonObj: Object): boolean => {
  const allowedProps: JsonAllowedProps = Object.keys(new JsonProps()) as JsonAllowedProps;
  return Object.keys(jsonObj).every(prop => allowedProps.includes(prop as keyof JsonProps));
};

export const createLinkableItems = (
  records: Entity[],
  recordIds: string[],
  dataverseService: IDataverseService): Entity[] => {
  records.forEach(record => {
    recordIds.push(record.id);
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

export const parseRawInput = (
  appWrapperProps: IAppWrapperProps,
  setIsLoading: React.Dispatch<React.SetStateAction<boolean>>,
  setError: React.Dispatch<React.SetStateAction<Error | undefined>>,
) => {
  const { fetchXmlOrJson } = appWrapperProps;

  const props: IFetchToSubgridProps = {
    error: undefined,
    _service: appWrapperProps._service,
    allocatedWidth: appWrapperProps.allocatedWidth,
    fetchXml: appWrapperProps.default.fetchXml,
    pageSize: getPageSize(appWrapperProps.default.pageSize),
    newButtonVisibility: appWrapperProps.default.newButtonVisibility,
    deleteButtonVisibility: appWrapperProps.default.deleteButtonVisibility,
    setIsLoading,
    setError,
    fetchXmlOrJson: null,
    default: {
      fetchXml: null,
      pageSize: 0,
      deleteButtonVisibility: false,
      newButtonVisibility: false,
    },
  };

  try {
    const fieldValueJson: IJsonProps = JSON.parse(fetchXmlOrJson ?? '') as IJsonProps;
    if (!isJsonValid(fieldValueJson)) props.error = new Error('JSON is not valid');

    if (fieldValueJson.fetchXml) props.fetchXml = fieldValueJson.fetchXml;
    if (fieldValueJson.pageSize) props.pageSize = fieldValueJson.pageSize;

    if (fieldValueJson.newButtonVisibility) {
      props.newButtonVisibility = fieldValueJson.newButtonVisibility;
    }

    if (fieldValueJson.deleteButtonVisibility) {
      props.deleteButtonVisibility = fieldValueJson.deleteButtonVisibility;
    }
  }
  catch {
    const fieldValueFetchXml: string | null = fetchXmlOrJson || appWrapperProps.default.fetchXml;
    const fetchXmlParserError: string | null = getFetchXmlParserError(fieldValueFetchXml);

    if (fetchXmlParserError) props.error = new Error(fetchXmlParserError);
    if (fieldValueFetchXml) props.fetchXml = fieldValueFetchXml;
  }

  return props as IFetchToSubgridProps;
};
