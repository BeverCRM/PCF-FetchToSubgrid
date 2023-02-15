import * as React from 'react';
import {
  ColumnActionsMode,
  ContextualMenu,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsFooterProps,
  IDetailsListProps,
  Stack,
  Selection,
} from '@fluentui/react';
import { LinkableItem } from './LinkableItems';
import {
  getColumns,
  getRecordsCount,
  openRecord,
  getEntityDisplayName } from '../services/crmService';
import { Footer } from './Footer';
import {
  getEntityName,
  getItems,
  isAggregate,
  getOrderInFetch } from '../utilities/utilities';
import { Loader } from './Loader';
import { InfoMessage } from './InfoMessage';
import { dataSetStyles } from '../styles/comandBarStyles';
import { CommandBar } from './ComandBar';
import {
  getContextualMenuProps,
  onColumnClick,
  onDialogClick,
  selectionChanged } from '../utilities/sortItems';

export interface IFetchSubgridProps {
  fetchXml: string | null;
  defaultPageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
  userParameters: any;
}

export const FetchSubgrid: React.FC<IFetchSubgridProps> = props => {
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
  const [currentPage, setCurrentPage] = React.useState(1);
  const [isDialogAccepted, setDialogAccepted] = React.useState(false);

  const recordIds = React.useRef<string[]>([]);
  const nextButtonDisable = React.useRef(true);
  const displayName = React.useRef('');
  const deleteBtnClassName = React.useRef('disableButton');
  const totalRecords = React.useRef(0);
  const selectItemsCount = React.useRef(0);
  const firstNumber = React.useRef(0);
  const lastNumber = React.useRef(0);
  const errorMessage = React.useRef<any>('');

  const isDeleteBtnVisible = userParameters?.DeleteButtonVisibility || deleteButtonVisibility;
  const isNewBtnVisible = userParameters?.NewButtonVisibility || newButtonVisibility;
  const pageSize: number = defaultPageSize;
 
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
    [currentPage, nextButtonDisable, selectItemsCount],
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
      setCurrentPage(1);
      try {
        displayName.current = await getEntityDisplayName(getEntityName(fetchXml ?? ''));
        const columns: IColumn[] = await getColumns(fetchXml);
        const order = getOrderInFetch(fetchXml ?? '');

        if (order) {
          const filteredColumns = columns.map(col => {
            if (col.ariaLabel === Object.values(order)[0] ||
            col.fieldName === Object.values(order)[0]) {

              col.isSorted = true;
              col.isSortedDescending = Object.keys(order)[0] === 'true';
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
  }, [fetchXml, userParameters]);

  React.useEffect(() => {
    deleteBtnClassName.current = 'disableButton';
    (async () => {
      setIsLoading(true);
      if (isDialogAccepted) return;

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
        lastNumber.current = (currentPage - 1) * pageSize + records.length;
        firstNumber.current = (currentPage - 1) * pageSize + 1;
        setItems(records);
      }
      catch (err) {
        console.log('Error', err);
        errorMessage.current = err;
        setColumns([]);
      }
      setIsLoading(false);
    })();
  }, [props, currentPage, isDialogAccepted]);

  if (isLoading) {
    return <Loader/>;
  }

  if (columns.length === 0) {
    return <InfoMessage message={errorMessage.current.message}/>;
  }

  const onContextualMenuDismissed = (): void => {
    setMenuProps({
      contextualMenuProps: undefined,
    });
  };

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

  const selection = new Selection({
    onSelectionChanged: () => {
      selectionChanged(
        selection,
        selectItemsCount,
        fetchXml,
        deleteBtnClassName,
        setSelectedRecordIds);
    },
  });

  return (
    <div className='fetchSubgridControl'>
      <Stack horizontal horizontalAlign="end" className={dataSetStyles.buttons}>
        <CommandBar
          entityName={getEntityName(fetchXml ?? '')}
          selectedRecordIds={selectedRecordIds}
          displayName={displayName.current}
          setDialogAccepted={setDialogAccepted}
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
