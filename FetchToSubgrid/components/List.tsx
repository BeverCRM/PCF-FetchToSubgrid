import * as React from 'react';
import { openRecord } from '../services/dataverseService';
import { Footer } from './Footer';
import { Entity } from '../utilities/types';
import { isAggregate } from '../utilities/fetchXmlUtils';
import {
  ColumnActionsMode,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IContextualMenuProps,
  IDetailsFooterProps,
  IDetailsListProps,
  IObjectWithKey,
  Selection } from '@fluentui/react';
import {
  getContextualMenuProps,
  onColumnClick,
  onDialogClick,
  selectionChanged } from '../utilities/sortingUtils';

interface IListProps {
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
  setItems: (items: Entity[]) => void;
  setColumns: (columns: IColumn[]) => void;
  setCurrentPage: (currentPage: number) => void;
  setSelectedRecordIds: React.Dispatch<React.SetStateAction<string[]>>
  setContextualMenuProps: React.Dispatch<React.SetStateAction<IContextualMenuProps | undefined>>;
}

export const List: React.FC<IListProps> = props => {
  const {
    entityName,
    recordIds,
    fetchXml,
    columns,
    items,
    deleteBtnClassName,
    pageSize,
    currentPage,
    nextButtonDisabled,
    firstItemIndex,
    lastItemIndex,
    selectedItemsCount,
    totalRecordsCount,
    setCurrentPage,
    setItems,
    setColumns,
    setSelectedRecordIds,
    setContextualMenuProps,
  } = props;

  const selection: Selection<IObjectWithKey> = new Selection({
    onSelectionChanged: () => {
      selectionChanged(
        selection,
        selectedItemsCount,
        fetchXml,
        deleteBtnClassName,
        setSelectedRecordIds);
    },
  });

  const onItemInvoked = React.useCallback((
    record: Entity,
    index?: number | undefined) : void => {
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    if (index !== undefined && !hasAggregate) {
      openRecord(entityName, recordIds.current[index]);
    }
  }, [fetchXml]);

  const onColumnContextMenu = (
    column?: IColumn,
    ev?: React.MouseEvent<HTMLElement, MouseEvent>): void => {
    if (column?.columnActionsMode !== ColumnActionsMode.disabled) {
      const onDialogClickParameters = {
        fetchXml,
        pageSize,
        currentPage,
        setItems,
        setColumns,
        recordIds,
        columns,
      };

      const onContextualMenuDismissed = (): void => setContextualMenuProps(undefined);

      setContextualMenuProps(getContextualMenuProps(
        ev,
        column,
        onContextualMenuDismissed,
        onDialogClick,
        onDialogClickParameters,
      ));
    }
  };

  const columnClick = (ev?: React.MouseEvent<HTMLElement, MouseEvent>, col?: IColumn) => {
    onColumnClick(ev, col, columns, onColumnContextMenu, setColumns);
  };

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = React.useCallback(
    (props?: IDetailsFooterProps) => {
      const movePreviousIsDisabled = currentPage <= 1;
      if (props) {
        return (
          <Footer
            firstItemIndex={firstItemIndex.current}
            lastItemIndex={lastItemIndex.current}
            selectedItems = {selectedItemsCount.current}
            totalRecordsCount={totalRecordsCount.current}
            currentPage={currentPage}
            setCurrentPage={setCurrentPage}
            nextButtonDisable={nextButtonDisabled.current}
            movePreviousIsDisabled={movePreviousIsDisabled}
          />
        );
      }
      return null;
    },
    [currentPage, nextButtonDisabled, selectedItemsCount],
  );

  return <DetailsList
    columns={columns}
    items={items}
    layoutMode={DetailsListLayoutMode.fixedColumns}
    onItemInvoked={onItemInvoked}
    onRenderDetailsFooter={onRenderDetailsFooter}
    onColumnHeaderClick={columnClick}
    selection={selection}
  />;
};
