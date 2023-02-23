import * as React from 'react';
import { openRecord } from '../services/dataverseService';
import { Footer } from './Footer';
import { Entity, IListProps } from '../utilities/types';
import { addOrderToFetch, isAggregate } from '../utilities/fetchXmlUtils';
import { filterColumns, getFilteredRecords } from '../utilities/utils';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsFooterProps,
  IDetailsListProps,
  IObjectWithKey,
  Selection,
} from '@fluentui/react';
import { LinkableItem } from './LinkableItems';

export const List: React.FC<IListProps> = props => {
  const {
    entityName,
    recordIds,
    fetchXml,
    columns,
    pageSize,
    items,
    deleteBtnClassName,
    currentPage,
    nextButtonDisabled,
    firstItemIndex,
    lastItemIndex,
    selectedItemsCount,
    totalRecordsCount,
    setCurrentPage,
    setColumns,
    setItems,
    setSelectedRecordIds,
  } = props;

  const onItemInvoked = React.useCallback((record?: Entity, index?: number | undefined): void => {
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    if (index !== undefined && !hasAggregate) {
      openRecord(entityName, recordIds.current[index]);
    }
  }, [fetchXml]);

  const onColumnHeaderClick = async (
    dialogEvent?: React.MouseEvent<HTMLElement, MouseEvent>,
    column?: IColumn) => {
    const fieldName = column?.className === 'linkEntity' ? column?.ariaLabel : column?.fieldName;

    const newFetchXml = addOrderToFetch(
      fetchXml ?? '',
      fieldName ?? '',
      dialogEvent,
      column);

    const filteredColumns: IColumn[] = filterColumns(
      column?.fieldName,
      column?.ariaLabel,
      undefined,
      columns) ?? [];

    const filteredRecords = await getFilteredRecords(
      totalRecordsCount,
      newFetchXml,
      pageSize,
      currentPage,
      nextButtonDisabled,
      lastItemIndex,
      firstItemIndex);

    filteredRecords.forEach(record => {
      recordIds.current.push(record.id);
      Object.keys(record).forEach(key => {
        if (key !== 'id') {
          const value: any = record[key];
          record[key] = value.isLinkable ? <LinkableItem item={value} /> : value.displayName;
        }
      });
    });

    setColumns(filteredColumns);
    setItems(filteredRecords);
  };

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = React.useCallback(
    (props?: IDetailsFooterProps) => {
      const movePreviousIsDisabled = currentPage <= 1;
      if (!props) return null;

      return (
        <Footer
          firstItemIndex={firstItemIndex.current}
          lastItemIndex={lastItemIndex.current}
          selectedItems={selectedItemsCount.current}
          totalRecordsCount={totalRecordsCount.current}
          currentPage={currentPage}
          setCurrentPage={setCurrentPage}
          nextButtonDisable={nextButtonDisabled.current}
          movePreviousIsDisabled={movePreviousIsDisabled}
        />
      );
    },
    [currentPage, nextButtonDisabled, selectedItemsCount],
  );

  const selectionChangeHandler = (selection: Selection<IObjectWithKey>) => {
    const currentSelection: IObjectWithKey[] = selection.getSelection();
    selectedItemsCount.current = currentSelection.length;

    deleteBtnClassName.current = currentSelection.length && !isAggregate(fetchXml ?? '')
      ? 'ms-Button'
      : 'disableButton';

    const recordIds: string[] = currentSelection.map((record: any) => record.id);
    setSelectedRecordIds(recordIds);
  };

  const selection = new Selection({
    onSelectionChanged: (): void => selectionChangeHandler(selection),
  });

  return <DetailsList
    columns={columns}
    items={items}
    layoutMode={DetailsListLayoutMode.fixedColumns}
    onItemInvoked={onItemInvoked}
    onRenderDetailsFooter={onRenderDetailsFooter}
    onColumnHeaderClick={onColumnHeaderClick}
    selection={selection}
  />;
};
