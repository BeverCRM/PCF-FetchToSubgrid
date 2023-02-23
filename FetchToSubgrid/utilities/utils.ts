import * as React from 'react';
import { IColumn } from '@fluentui/react';
import { getRecordsCount } from '../services/dataverseService';
import { getColumns, getItems } from './d365Utils';
import { AttributeType } from './enums';
import { getOrderInFetch } from './fetchXmlUtils';
import { Entity } from './types';

export const checkIfAttributeIsEntityReferance = (attributeType: AttributeType): boolean =>
  attributeType === AttributeType.Lookup ||
  attributeType === AttributeType.Owner ||
  attributeType === AttributeType.Customer;

export const needToGetFormattedValue = (attributeType: AttributeType) =>
  attributeType === AttributeType.Money ||
  attributeType === AttributeType.PickList ||
  attributeType === AttributeType.DateTime ||
  attributeType === AttributeType.MultiselectPickList ||
  attributeType === AttributeType.TwoOptions;

export const filterColumns = (
  fieldName?: string,
  ariaLabel?: string,
  descending? : boolean,
  allColumns?: IColumn[]): IColumn[] | undefined => {
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

  return filteredColumns;
};

export const getFilteredColumns = async (fetchXml: string | null): Promise<IColumn[]> => {
  const columns: IColumn[] = await getColumns(fetchXml);
  const order = getOrderInFetch(fetchXml ?? '');

  if (!order) return columns;

  const filteredColumns: IColumn[] = filterColumns(
    Object.keys(order)[0],
    Object.keys(order)[0],
    Object.values(order)[0],
    columns) ?? [];

  return filteredColumns;
};

export const getFilteredRecords = async (
  totalRecordsCount: React.MutableRefObject<number>,
  fetchXml: string | null,
  pageSize: number,
  currentPage: number,
  nextButtonDisabled: React.MutableRefObject<boolean>,
  lastItemIndex: React.MutableRefObject<number>,
  firstItemIndex: React.MutableRefObject<number>): Promise<Entity[]> => {
  const recordsCount: number = await getRecordsCount(fetchXml ?? '');

  totalRecordsCount.current = recordsCount;
  nextButtonDisabled.current = Math.ceil(recordsCount / pageSize) <= currentPage;

  const records: Entity[] = await getItems(
    fetchXml,
    pageSize,
    currentPage,
    recordsCount);

  lastItemIndex.current = (currentPage - 1) * pageSize + records.length;
  firstItemIndex.current = (currentPage - 1) * pageSize + 1;

  return records;
};
