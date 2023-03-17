import * as React from 'react';
import { IColumn } from '@fluentui/react';
import { getColumns } from './d365Utils';
import { AttributeType } from '../@types/enum';
import { getOrderInFetchXml } from './fetchXmlUtils';
import { LinkableItem } from '../components/LinkableItems';
import {
  Entity,
  IDataverseService,
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
