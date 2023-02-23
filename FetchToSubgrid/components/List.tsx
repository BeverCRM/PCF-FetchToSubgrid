import * as React from 'react';
import { Entity, IListProps } from '../utilities/types';
import { addOrderToFetch, isAggregate } from '../utilities/fetchXmlUtils';
import { calculateFilteredRecordsData, sortColumns } from '../utilities/utils';
import { Footer } from './Footer';
import { LinkableItem } from './LinkableItems';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsFooterProps,
  IDetailsListProps,
  IObjectWithKey,
  Selection,
} from '@fluentui/react';
import { getItems } from '../utilities/d365Utils';

export const List: React.FC<IListProps> = props => {
  const {
    _service: dataverseService,
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
      dataverseService.openRecord(entityName, recordIds.current[index]);
    }
  }, [fetchXml]);

  const onColumnHeaderClick = async (
    dialogEvent?: React.MouseEvent<HTMLElement, MouseEvent>,
    column?: IColumn): Promise<void> => {
    const fieldName = column?.className === 'linkEntity' ? column?.ariaLabel : column?.fieldName;

    const newFetchXml = addOrderToFetch(
      fetchXml ?? '',
      fieldName ?? '',
      dialogEvent,
      column);

    const filteredColumns: IColumn[] = sortColumns(
      column?.fieldName,
      column?.ariaLabel,
      undefined,
      columns) ?? [];

    const recordsCount: number = await dataverseService.getRecordsCount(newFetchXml ?? '');
    const filteredRecords: Entity[] = await getItems(
      newFetchXml,
      pageSize,
      currentPage,
      recordsCount,
      dataverseService);

    calculateFilteredRecordsData(
      totalRecordsCount,
      recordsCount,
      filteredRecords,
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

          // eslint-disable-next-line no-extra-parens
          record[key] = value.isLinkable ? (
            <LinkableItem
              _service={dataverseService}
              item={value}
            />
          ) : value.displayName;
        }
      });
    });

    setColumns(filteredColumns);
    setItems(filteredRecords);
  };

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = React.useCallback(
    (props?: IDetailsFooterProps): JSX.Element | null => {
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
