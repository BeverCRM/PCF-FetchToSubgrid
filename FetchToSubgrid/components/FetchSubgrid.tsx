import * as React from 'react';
import { ContextualMenu, IColumn, Stack } from '@fluentui/react';
import { getColumns, getEntityDisplayName } from '../services/crmService';
import { dataSetStyles } from '../styles/comandBarStyles';
import { LinkableItem } from './LinkableItems';
import { CommandBar } from './ComandBar';
import { List } from './List';
import { Entity } from '../utilities/types';
import { getEntityName } from '../utilities/fetchXmlUtils';
import { getFilteredRecords, setFilteredColumns } from '../utilities/d365Utils';

export interface IFetchSubgridProps {
  fetchXml: string | null;
  defaultPageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
  setIsLoading: (isLoading: boolean) => void;
  setErrorMessage: (message?: string) => void;
  isVisible: boolean;
}

export const FetchSubgrid: React.FC<IFetchSubgridProps> = props => {
  const {
    deleteButtonVisibility,
    newButtonVisibility,
    defaultPageSize,
    fetchXml,
    setErrorMessage,
    setIsLoading,
    isVisible,
  } = props;

  const [items, setItems] = React.useState<Entity[]>([]);
  const [menuProps, setMenuProps] = React.useState<any>();
  const [selectedRecordIds, setSelectedRecordIds] = React.useState<string[]>([]);
  const [currentPage, setCurrentPage] = React.useState(1);
  const [isDialogAccepted, setDialogAccepted] = React.useState(false);
  const [columns, setColumns] = React.useState<IColumn[]>([]);

  const recordIds = React.useRef<string[]>([]);
  const nextButtonDisabled = React.useRef(true);
  const displayName = React.useRef('');
  const deleteBtnClassName = React.useRef('disableButton');
  const totalRecordsCount = React.useRef(0);
  const selectedItemsCount = React.useRef(0);
  const firstItemIndex = React.useRef(0);
  const lastItemIndex = React.useRef(0);

  const pageSize: number = defaultPageSize;
  const entityName = getEntityName(fetchXml ?? '');

  React.useEffect(() => {
    (async () => {
      try {
        setCurrentPage(1);
        setFilteredColumns(fetchXml, setColumns, getColumns);
        displayName.current = await getEntityDisplayName(entityName);
      }
      catch (err: any) {
        setErrorMessage(err.message);
      }
    })();
  }, [fetchXml, deleteButtonVisibility, newButtonVisibility, pageSize]);

  React.useEffect(() => {
    deleteBtnClassName.current = 'disableButton';
    (async () => {
      if (isDialogAccepted) return;
      setIsLoading(true);
      setErrorMessage();
      try {
        const records = await getFilteredRecords(
          totalRecordsCount,
          fetchXml,
          pageSize,
          currentPage,
          nextButtonDisabled,
          lastItemIndex,
          firstItemIndex);

        records.forEach(record => {
          recordIds.current.push(record.id);
          Object.keys(record).forEach(key => {
            if (key !== 'id') {
              const value: any = record[key];
              record[key] = value.linkable ? <LinkableItem item={value} /> : value.displayName;
            }
          });
        });
        setItems(records);
      }
      catch (err: any) {
        setErrorMessage(err.message);
      }
      setIsLoading(false);
    })();
  }, [fetchXml, pageSize, currentPage, isDialogAccepted]);

  return (
    <div className='fetchSubgridControl' style={{ display: isVisible ? 'grid' : 'none' }}>
      <Stack horizontal horizontalAlign="end" className={dataSetStyles.buttons}>
        <CommandBar
          entityName={entityName}
          selectedRecordIds={selectedRecordIds}
          displayName={displayName.current}
          setDialogAccepted={setDialogAccepted}
          className={deleteBtnClassName.current}
          deleteButtonVisibility={deleteButtonVisibility}
          newButtonVisibility={newButtonVisibility}
        />
      </Stack>

      <List entityName={entityName}
        recordIds={recordIds}
        fetchXml={fetchXml}
        columns={columns}
        items={items}
        deleteBtnClassName={deleteBtnClassName}
        setSelectedRecordIds={setSelectedRecordIds}
        pageSize={pageSize}
        currentPage={currentPage}
        setItems={setItems}
        setColumns={setColumns}
        setMenuProps={setMenuProps}
        firstItemIndex={firstItemIndex}
        lastItemIndex={lastItemIndex}
        selectedItemsCount={selectedItemsCount}
        totalRecordsCount={totalRecordsCount}
        setCurrentPage={setCurrentPage}
        nextButtonDisabled={nextButtonDisabled}
      />
      <ContextualMenu {...menuProps.contextualMenuProps} />
    </div>
  );
};
