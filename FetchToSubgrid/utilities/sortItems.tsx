import {
  DirectionalHint,
  IColumn,
  IContextualMenuItem,
  IObjectWithKey } from '@fluentui/react';
import * as React from 'react';
import { LinkableItem } from '../components/LinkableItems';
import { getRecordsCount } from '../services/crmService';
import { addOrderToFetch, getItems, isAggregate } from './utilities';

export const onDialogClick = async (
  column?: IColumn,
  dialogEvent?: React.MouseEvent<HTMLElement, MouseEvent>,
  onDialogClickParameters? : any,
) => {
  const {
    fetchXml,
    pageSize,
    currentPage,
    setItems,
    setColumns,
    recordIds,
    columns,
  } = onDialogClickParameters;

  const fieldName = column?.className === 'linkEntity' ? column?.ariaLabel : column?.fieldName;

  const newFetchXml = addOrderToFetch(
    fetchXml ?? '',
    fieldName ?? '',
    dialogEvent,
    column,
  );

  const recordsCount: number = await getRecordsCount(fetchXml ?? '');
  const records: ComponentFramework.WebApi.Entity[] = await getItems(
    newFetchXml,
    pageSize,
    currentPage,
    recordsCount);

  records.forEach(record => {
    recordIds.current.push(record.id);

    Object.keys(record).forEach(key => {
      if (key !== 'id') {
        const value: any = record[key];
        record[key] = value.linkable ? <LinkableItem item = {value} /> : value.displayName;
      }
    });
  });

  setItems(records);

  const filteredColumns = columns.map((col: IColumn) => {
    if (column?.key === col.key) {
      // @ts-ignore
      col.isSortedDescending = dialogEvent?.key === 'zToA';
      col.isSorted = true;
    }
    else {
      col.isSorted = false;
    }
    return col;
  });

  setColumns(filteredColumns);
};

export const getContextualMenuProps = (
  ev?: React.MouseEvent<HTMLElement, MouseEvent>,
  column?: IColumn,
  onContextualMenuDismissed?: any,
  onDialogClick?: any,
  onDialogClickParameters?: any,

): Object => {

  const items: IContextualMenuItem[] = [
    {
      key: 'aToZ',
      name: 'Sort A to Z',
      iconProps: { iconName: 'SortUp' },
      checked: true,
      canCheck:  column?.isSorted && !column?.isSortedDescending,
    },
    {
      key: 'zToA',
      name: 'Sort Z to A',
      iconProps: { iconName: 'SortDown' },
      checked: true,
      canCheck: column?.isSorted && column?.isSortedDescending,
    },
  ];

  return {
    items,
    target: ev?.currentTarget as HTMLElement,
    directionalHint: DirectionalHint.bottomLeftEdge,
    gapSpace: 10,
    isBeakVisible: true,
    onDismiss: onContextualMenuDismissed,
    onItemClick: (col: IColumn, dialogEvent: React.MouseEvent<HTMLElement>) => {
      onDialogClick(column, dialogEvent, onDialogClickParameters);
    },
  };
};

export const selectionChanged = (selection: any,
  selectItemsCount: any,
  fetchXml: any,
  deleteBtnClassName: any,
  setSelectedRecordIds:any) => {
  const currentSelection: IObjectWithKey[] = selection.getSelection();
  selectItemsCount.current = currentSelection.length;

  currentSelection.length && !isAggregate(fetchXml ?? '')
    ? deleteBtnClassName.current = 'ms-Button'
    : deleteBtnClassName.current = 'disableButton';
  const recordIds: string[] = currentSelection.map((record: any) => record.id);
  setSelectedRecordIds(recordIds);
};

export const onColumnClick = (
  ev: any,
  col?: IColumn,
  columns?: any,
  onColumnContextMenu?: any,
  setColumns?: any,
) => {
  const key = col?.key;
  const filteredColumns = columns.map((column: IColumn) => {
    if (column.key === key) {
      onColumnContextMenu(col, ev);
      column.onColumnClick = (ev, columnContex) => {
        onColumnContextMenu(columnContex, ev);
      };
    }
    else {
      column.isSorted = false;
    }
    return column;
  });

  setColumns(filteredColumns);
};
