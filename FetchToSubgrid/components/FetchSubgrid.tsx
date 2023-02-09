import * as React from 'react';
import {
  ColumnActionsMode,
  ContextualMenu,
  DetailsList,
  DetailsListLayoutMode,
  DirectionalHint,
  IColumn,
  IContextualMenuItem,
  IDetailsFooterProps,
  IDetailsListProps,
  Stack,
  Selection,
  IObjectWithKey,
} from '@fluentui/react';
import { LinkableItem } from './LinkableItems';
import {
  getPagingLimit,
  getColumns,
  getRecordsCount,
  openRecord,
  getEntityDisplayName } from '../services/crmService';
import { Footer } from './Footer';
import {
  getCountInFetchXml,
  getEntityName,
  getItems,
  getPageInFetch,
  isAggregate,
  addOrderToFetch,
  getOrderInFetch } from '../utilities/utilities';
import { Loader } from './Loader';
import { InfoMessage } from './InfoMessage';
import { dataSetStyles } from '../styles/comandBarStyles';
import { CommandBar } from './ComandBar';

export interface IFetchSubgridProps {
  fetchXml: string | null;
  defaultPageSize: number | null;
  deleteButtonVisibility: string | null;
  newButtonVisibility: string | null;
  userParameters: any;
}

export const FetchSubgrid: React.FunctionComponent<IFetchSubgridProps> = props => {
  const {
    userParameters,
    deleteButtonVisibility,
    newButtonVisibility,
    defaultPageSize,
    fetchXml } = props;

  const [isLoading, setIsLoading] = React.useState(false);
  const [columns, setColumns] = React.useState<IColumn[]>([]);
  const [Items, setItems] = React.useState<ComponentFramework.WebApi.Entity[]>([]);
  const [menuProps, setMenuProps] = React.useState<any>({ contextualMenuProps: undefined });
  const [selectedRecordIds, setSelectedRecordIds] = React.useState<any>([]);
  const [currentPage, setCurrentPage] = React.useState(getPageInFetch(fetchXml));
  const [isUsedButton, setIsUsedButton] = React.useState(false);

  const recordIds = React.useRef<string[]>([]);
  const nextButtonDisable = React.useRef(true);
  const displayName = React.useRef('');
  const deleteBtnClassName = React.useRef('disableButton');
  const totalRecords = React.useRef(0);
  const selectItemsCount = React.useRef(0);
  const firstNumber = React.useRef(0);
  const lastNumber = React.useRef(0);

  const isDeleteBtnVisible = userParameters?.DeleteButtonVisibility || deleteButtonVisibility;
  const isNewBtnVisible = userParameters?.NewButtonVisibility || newButtonVisibility;

  const pageSize: number = getCountInFetchXml(fetchXml) || defaultPageSize || getPagingLimit();

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = React.useCallback(
    (props: IDetailsFooterProps | undefined) => {
      const isMovePrevious = !(currentPage > 1);
      if (props) {
        return (
          <Footer
            firstNumber={firstNumber.current}
            lastNumber={lastNumber.current}
            selectedItems = {selectItemsCount.current}
            totalRecordsCount={totalRecords.current}
            currentPage={currentPage}
            setCurrentPage={setCurrentPage}
            nextButtonDisable={nextButtonDisable.current}
            isMovePrevious={isMovePrevious}
          />);
      }
      return null;
    },
    [currentPage, nextButtonDisable],
  );

  const onItemInvoked = React.useCallback((
    record: ComponentFramework.WebApi.Entity,
    index?: number | undefined) : void => {
    const entityName: string = getEntityName(fetchXml ?? '');
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');

    if (index !== undefined && !hasAggregate) {
      openRecord(entityName, recordIds.current[index]);
    }
  }, [Items]);

  React.useEffect(() => {
    (async () => {
      setCurrentPage(getPageInFetch(fetchXml));
      try {
        displayName.current = await getEntityDisplayName(getEntityName(fetchXml ?? ''));
        const columns: IColumn[] = await getColumns(fetchXml);

        const order = getOrderInFetch(fetchXml ?? '');
        if (order) {
          const filteredColumns = columns.map(col => {
            if (col.fieldName === Object.values(order)[0]) {
              col.isSorted = true;
              col.isSortedDescending = !(Object.keys(order)[0] === 'true');
              return col;
            }
            return col;
          });

          setColumns(filteredColumns);
        }
        else {
          setColumns(columns);
        }
      }
      catch {
        setColumns([]);
      }
    })();
  }, [fetchXml]);

  React.useEffect(() => {
    deleteBtnClassName.current = 'disableButton';
    (async () => {
      setIsLoading(true);
      try {
        const recordsCount: number = await getRecordsCount(fetchXml ?? '');
        totalRecords.current = recordsCount;
        if (Math.ceil(recordsCount / pageSize) > currentPage) {
          nextButtonDisable.current = false;
        }
        else {
          nextButtonDisable.current = true;
        }

        const records: ComponentFramework.WebApi.Entity[] = await getItems(
          fetchXml,
          pageSize,
          currentPage);

        records.forEach(record => {
          recordIds.current.push(record.id);

          Object.keys(record).forEach(key => {
            if (key !== 'id') {
              const value: any = record[key];
              record[key] = value.linkable ? <LinkableItem item = {value} /> : value.displayName;
            }
          });
        });
        lastNumber.current = (currentPage - 1) * pageSize + records.length;
        firstNumber.current = (currentPage - 1) * pageSize + 1;
        setItems(records);
      }
      catch (err) {
        console.log('Error', err);
      }
      setIsLoading(false);
    })();
  }, [props, currentPage, isUsedButton]);

  if (isLoading) {
    return <Loader/>;
  }

  if (columns.length === 0) {
    return <InfoMessage fetchXml ={fetchXml}/>;
  }

  const onContextualMenuDismissed = (): void => {
    setMenuProps({
      contextualMenuProps: undefined,
    });
  };

  const onDialogClick = async (column?: IColumn,
    dialogEvent?: React.MouseEvent<HTMLElement, MouseEvent>) => {
    const fieldName = column?.fieldName;
    const newFetchXml = addOrderToFetch(fetchXml ?? '', fieldName ?? '', dialogEvent);

    const records: ComponentFramework.WebApi.Entity[] = await getItems(
      newFetchXml,
      pageSize,
      currentPage);

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

    const filteredColumns = columns.map(col => {
      if (column?.key === col.key) {
        // @ts-ignore
        col.isSortedDescending = dialogEvent?.checked;
        col.isSorted = true;
      }
      else {
        col.isSorted = false;
      }
      return col;
    });

    setColumns(filteredColumns);
  };

  const getContextualMenuProps = (ev?: React.MouseEvent<HTMLElement, MouseEvent> | undefined,
    column?: IColumn): Object => {
    const items: IContextualMenuItem[] = [
      {
        key: 'aToZ',
        name: 'A to Z',
        iconProps: { iconName: 'SortUp' },
        checked: false,
      },
      {
        key: 'zToA',
        name: 'Z to A',
        iconProps: { iconName: 'SortDown' },
        checked: true,
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
        onDialogClick(column, dialogEvent);
      },
    };
  };

  const onColumnContextMenu = (column?: IColumn,
    ev?: React.MouseEvent<HTMLElement, MouseEvent> | undefined): void => {
    if (column?.columnActionsMode !== ColumnActionsMode.disabled) {
      setMenuProps({
        contextualMenuProps: getContextualMenuProps(ev, column),
      });
    }
  };

  const columnClick = (ev?: React.MouseEvent<HTMLElement, MouseEvent> | undefined,
    col?: IColumn) => {
    const key = col?.key;
    const filteredColumns = columns.map(column => {
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

  const selection = new Selection({
    onSelectionChanged: () => {
      const currentSelection: IObjectWithKey[] = selection.getSelection();
      selectItemsCount.current = currentSelection.length;

      currentSelection.length && !isAggregate(fetchXml ?? '')
        ? deleteBtnClassName.current = 'ms-Button'
        : deleteBtnClassName.current = 'disableButton';
      const recordIds: string[] = currentSelection.map((record: any) => record.id);
      setSelectedRecordIds(recordIds);
    },
  });

  return (
    <div className='fetchSubgridControl'>
      <Stack horizontal horizontalAlign="end" className={dataSetStyles.buttons}>
        <CommandBar
          entityName={getEntityName(fetchXml ?? '')}
          selectedRecordIds={selectedRecordIds}
          displayName={displayName.current}
          setIsUsedButton={setIsUsedButton}
          className = {deleteBtnClassName.current}
          deleteButtonVisibility={isDeleteBtnVisible}
          newButtonVisibility={isNewBtnVisible}
        />
      </Stack>
      <DetailsList
        columns={columns}
        items={Items}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        onItemInvoked={onItemInvoked}
        onRenderDetailsFooter={onRenderDetailsFooter}
        onColumnHeaderClick={columnClick}
        selection={selection}
      />
      {<ContextualMenu {...menuProps.contextualMenuProps} />}
    </div>
  );
};
