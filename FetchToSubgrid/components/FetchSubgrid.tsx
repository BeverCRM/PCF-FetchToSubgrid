import * as React from 'react';
import { IColumn, Stack } from '@fluentui/react';
import { getEntityDisplayName } from '../services/dataverseService';
import { dataSetStyles } from '../styles/comandBarStyles';
import { LinkableItem } from './LinkableItems';
import { CommandBar } from './ComandBar';
import { List } from './List';
import { Entity, IFetchSubgridProps } from '../utilities/types';
import { getEntityNameFromFetchXml } from '../utilities/fetchXmlUtils';
import { getFilteredRecords, getFilteredColumns } from '../utilities/utils';

export const FetchSubgrid: React.FC<IFetchSubgridProps> = props => {
  const {
    deleteButtonVisibility,
    newButtonVisibility,
    defaultPageSize,
    fetchXml,
    isVisible,
    setErrorMessage,
    setIsLoading,
  } = props;

  const [items, setItems] = React.useState<Entity[]>([]);
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
  const entityName = getEntityNameFromFetchXml(fetchXml ?? '');

  React.useEffect(() => {
    (async () => {
      try {
        setCurrentPage(1);
        const filteredColumns = await getFilteredColumns(fetchXml);
        setColumns(filteredColumns);
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
              record[key] = value.isLinkable ? <LinkableItem item={value} /> : value.displayName;
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
    <div className='FetchSubgridControl' style={{ display: isVisible ? 'grid' : 'none' }}>
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
        deleteBtnClassName={deleteBtnClassName}
        pageSize={pageSize}
        firstItemIndex={firstItemIndex}
        lastItemIndex={lastItemIndex}
        selectedItemsCount={selectedItemsCount}
        totalRecordsCount={totalRecordsCount}
        nextButtonDisabled={nextButtonDisabled}
        fetchXml={fetchXml}
        recordIds={recordIds}
        setSelectedRecordIds={setSelectedRecordIds}
        currentPage={currentPage}
        setCurrentPage={setCurrentPage}
        columns={columns}
        setColumns={setColumns}
        items={items}
        setItems={setItems}
      />
    </div>
  );
};
