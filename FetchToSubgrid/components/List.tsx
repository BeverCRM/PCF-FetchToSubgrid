import * as React from 'react';
import { Entity, IListProps } from '../@types/types';
import { addOrderToFetch, isAggregate } from '../utilities/fetchXmlUtils';
import { calculateFilteredRecordsData, createLinkableItems, sortColumns } from '../utilities/utils';
import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react';
import { getItems } from '../utilities/d365Utils';

export const List: React.FC<IListProps> = props => {
  const {
    _service: dataverseService,
    entityName,
    recordIds,
    fetchXml,
    columns,
    pageSize,
    forceReRender,
    items,
    currentPage,
    nextButtonDisabled,
    firstItemIndex,
    lastItemIndex,
    totalRecordsCount,
    selection,
    setColumns,
    setItems,
  } = props;

  const onItemInvoked = React.useCallback((record?: Entity, index?: number | undefined): void => {
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    if (index !== undefined && !hasAggregate) {
      dataverseService.openRecord(entityName, recordIds.current[index]);
    }
  }, [fetchXml]);

  const onColumnHeaderClick = React.useCallback(async (
    dialogEvent?: React.MouseEvent<HTMLElement, MouseEvent>,
    column?: IColumn): Promise<void> => {
    if (column?.className === 'colIsNotSortable') return;

    const fieldName = column?.className === 'linkEntity' ? column?.ariaLabel : column?.fieldName;

    const sortedColumns: IColumn[] = sortColumns(
      column?.fieldName,
      column?.ariaLabel,
      undefined,
      columns);

    const newFetchXml = addOrderToFetch(
      fetchXml ?? '',
      fieldName ?? '',
      column);

    const sortedRecords: Entity[] = await getItems(
      newFetchXml,
      pageSize,
      currentPage,
      totalRecordsCount,
      dataverseService);

    calculateFilteredRecordsData(
      totalRecordsCount,
      sortedRecords,
      pageSize,
      currentPage,
      nextButtonDisabled,
      lastItemIndex,
      firstItemIndex);

    const linkableItems = createLinkableItems(sortedRecords, recordIds.current, dataverseService);

    setColumns(sortedColumns);
    setItems(linkableItems);
  }, [
    columns,
    currentPage,
    dataverseService,
    fetchXml,
    firstItemIndex,
    nextButtonDisabled,
    pageSize,
    recordIds,
    totalRecordsCount,
  ]);

  return <DetailsList
    key={forceReRender}
    columns={columns}
    items={items}
    layoutMode={DetailsListLayoutMode.fixedColumns}
    onItemInvoked={onItemInvoked}
    onColumnHeaderClick={onColumnHeaderClick}
    selection={selection}
  />;
};
