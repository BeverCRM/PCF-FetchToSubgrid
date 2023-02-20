import {
  ColumnActionsMode,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsFooterProps,
  IDetailsListProps,
  IObjectWithKey,
  Selection } from '@fluentui/react';
import * as React from 'react';
import { openRecord } from '../services/crmService';
import { isAggregate } from '../utilities/utilities';
import { getContextualMenuProps, onColumnClick, onDialogClick } from '../utilities/sortItems';
import { Footer } from './Footer';
import { Entity } from '../utilities/types';

interface IListProps {
  entityName: string;
  recordIds: React.MutableRefObject<string[]>;
  fetchXml: string | null;
  columns: IColumn[];
  Items: Entity[];
  selection: Selection<IObjectWithKey>;
  pageSize: number;
  currentPage: number;
  setItems: (items: Entity[]) => void;
  setColumns: (columns: IColumn[]) => void;
  setMenuProps: any;
  firstItemIndex: React.MutableRefObject<number>;
  lastItemIndex: React.MutableRefObject<number>;
  selectedItemsCount: React.MutableRefObject<number>;
  totalRecordsCount: React.MutableRefObject<number>;
  setCurrentPage: (currentPage: number) => void;
  nextButtonDisabled: React.MutableRefObject<boolean>;
}

export const List: React.FC<IListProps> = props => {
  const {
    entityName,
    recordIds,
    fetchXml,
    columns,
    Items,
    selection,
    pageSize,
    currentPage,
    setItems,
    setColumns,
    setMenuProps,
    firstItemIndex,
    lastItemIndex,
    selectedItemsCount,
    totalRecordsCount,
    setCurrentPage,
    nextButtonDisabled,
  } = props;

  const onContextualMenuDismissed = (): void => {
    setMenuProps({
      contextualMenuProps: undefined,
    });
  };

  const onItemInvoked = React.useCallback((
    record: ComponentFramework.WebApi.Entity,
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

      setMenuProps({
        contextualMenuProps: getContextualMenuProps(
          ev,
          column,
          onContextualMenuDismissed,
          onDialogClick,
          onDialogClickParameters,
        ),
      });
    }
  };

  const columnClick = (ev?: React.MouseEvent<HTMLElement, MouseEvent>, col?: IColumn) => {
    onColumnClick(ev, col, columns, onColumnContextMenu, setColumns);
  };

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = React.useCallback(
    (props: IDetailsFooterProps | undefined) => {
      const movePreviousIsDisabled = !(currentPage > 1);
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
          />);
      }
      return null;
    },
    [currentPage, nextButtonDisabled, selectedItemsCount],
  );

  return <DetailsList
    columns={columns}
    items={Items}
    layoutMode={DetailsListLayoutMode.fixedColumns}
    onItemInvoked={onItemInvoked}
    onRenderDetailsFooter={onRenderDetailsFooter}
    onColumnHeaderClick={columnClick}
    selection={selection}
  />;
};
