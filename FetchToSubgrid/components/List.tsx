import * as React from 'react';
import { Entity, IService } from '../@types/types';
import { isAggregate } from '../utilities/fetchXmlUtils';
import { sortColumns } from '../utilities/utils';
import { DetailsList, DetailsListLayoutMode, IColumn, ISelection } from '@fluentui/react';
import { IDataverseService } from '../services/dataverseService';

interface IListProps extends IService<IDataverseService> {
  entityName: string;
  fetchXml: string | null;
  forceReRender: number;
  columns: IColumn[];
  items: Entity[];
  selection: ISelection;
  setSortingData: any;
}

export const List: React.FC<IListProps> = props => {
  const {
    _service: dataverseService,
    entityName,
    fetchXml,
    columns,
    forceReRender,
    items,
    selection,
    setSortingData,
  } = props;

  const onItemInvoked = React.useCallback((record?: Entity, index?: number | undefined): void => {
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    if (index !== undefined && !hasAggregate) {
      dataverseService.openRecordForm(entityName, record?.id);
    }
  }, [fetchXml]);

  const onColumnHeaderClick = React.useCallback(async (
    dialogEvent?: React.MouseEvent<HTMLElement, MouseEvent>,
    column?: IColumn): Promise<void> => {
    if (column?.className === 'colIsNotSortable') return;

    const fieldName = column?.className === 'linkEntity' ? column?.ariaLabel : column?.fieldName;
    sortColumns(column?.fieldName, column?.ariaLabel, undefined, columns);
    setSortingData({ fieldName, column });
  }, [columns, fetchXml]);

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
