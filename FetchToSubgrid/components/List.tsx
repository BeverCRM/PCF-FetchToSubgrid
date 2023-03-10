import * as React from 'react';
import { Entity, IListProps } from '../utilities/types';
import { addOrderToFetch, isAggregate } from '../utilities/fetchXmlUtils';
import { calculateFilteredRecordsData, sortColumns } from '../utilities/utils';
import { LinkableItem } from './LinkableItems';
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
    allocatedWidthKey,
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

  const updatedFetchXml = React.useRef(fetchXml);
  React.useEffect(() => { updatedFetchXml.current = fetchXml; }, [fetchXml]);

  const onItemInvoked = React.useCallback((record?: Entity, index?: number | undefined): void => {
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    if (index !== undefined && !hasAggregate) {
      dataverseService.openRecord(entityName, recordIds.current[index]);
    }
  }, [fetchXml]);

  const onColumnHeaderClick = async (
    dialogEvent?: React.MouseEvent<HTMLElement, MouseEvent>,
    column?: IColumn): Promise<void> => {
    if (column?.className === 'colIsNotSortable') return;

    const fieldName = column?.className === 'linkEntity' ? column?.ariaLabel : column?.fieldName;

    const sortedColumns: IColumn[] = sortColumns(
      column?.fieldName,
      column?.ariaLabel,
      undefined,
      columns) ?? [];

    const newFetchXml = addOrderToFetch(
      fetchXml ?? '',
      fieldName ?? '',
      column);

    const sortedRecords: Entity[] = await getItems(
      newFetchXml,
      pageSize,
      currentPage,
      totalRecordsCount.current,
      dataverseService);

    calculateFilteredRecordsData(
      totalRecordsCount.current,
      sortedRecords,
      pageSize,
      currentPage,
      nextButtonDisabled,
      lastItemIndex,
      firstItemIndex);

    sortedRecords.forEach(record => {
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

    setColumns(sortedColumns);
    setItems(sortedRecords);
  };

  return <DetailsList
    key={allocatedWidthKey}
    columns={columns}
    items={items}
    layoutMode={DetailsListLayoutMode.fixedColumns}
    onItemInvoked={onItemInvoked}
    onColumnHeaderClick={onColumnHeaderClick}
    selection={selection}
  />;
};
