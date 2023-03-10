import * as React from 'react';
import { IColumn } from '@fluentui/react';
import { getColumns } from './d365Utils';
import { AttributeType } from './enums';
import { getOrderInFetch } from './fetchXmlUtils';
import { Entity, IDataverseService } from './types';

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

export const sortColumns = (
  fieldName?: string,
  ariaLabel?: string,
  descending? : boolean,
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
  const order = getOrderInFetch(fetchXml ?? '');

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
