import * as React from 'react';
import { Entity, IService } from '../@types/types';
import { addOrderToFetch, isAggregate } from '../utilities/fetchXmlUtils';
import { createLinkableItems, sortColumns } from '../utilities/utils';
import { DetailsList, DetailsListLayoutMode, IColumn, ISelection } from '@fluentui/react';
import { getItems } from '../utilities/d365Utils';
import { IDataverseService } from '../services/dataverseService';

interface IListProps extends IService<IDataverseService> {
  entityName: string;
  fetchXml: string | null;
  pageSize: number;
  forceReRender: number;
  currentPage: number;
  recordIds: React.MutableRefObject<string[]>;
  columns: IColumn[];
  items: Entity[];
  selectedItemsCount: React.MutableRefObject<number>;
  totalRecordsCount: number;
  nextButtonDisabled: boolean;
  setItems: React.Dispatch<React.SetStateAction<ComponentFramework.WebApi.Entity[]>>;
  setColumns: React.Dispatch<React.SetStateAction<IColumn[]>>;
  setCurrentPage: React.Dispatch<React.SetStateAction<number>>;
  selection: ISelection;
}

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
    totalRecordsCount,
    selection,
    setColumns,
    setItems,
  } = props;

  const onItemInvoked = React.useCallback((record?: Entity, index?: number | undefined): void => {
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    if (index !== undefined && !hasAggregate) {
      dataverseService.openRecordForm(entityName, recordIds.current[index]);
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

    const linkableItems = createLinkableItems(sortedRecords, recordIds.current, dataverseService);

    setColumns(sortedColumns);
    setItems(linkableItems);
  }, [
    columns,
    currentPage,
    fetchXml,
    nextButtonDisabled,
    pageSize,
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
