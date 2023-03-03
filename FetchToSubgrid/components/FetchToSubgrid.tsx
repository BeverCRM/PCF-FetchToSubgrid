import * as React from 'react';
import { IColumn, Stack } from '@fluentui/react';
import { Entity, IFetchToSubgridProps } from '../utilities/types';
import { getSortedColumns, calculateFilteredRecordsData } from '../utilities/utils';
import { getEntityNameFromFetchXml } from '../utilities/fetchXmlUtils';
import { dataSetStyles } from '../styles/comandBarStyles';
import { LinkableItem } from './LinkableItems';
import { CommandBar } from './ComandBar';
import { List } from './List';
import { getItems } from '../utilities/d365Utils';

export const FetchToSubgrid: React.FC<IFetchToSubgridProps> = props => {
  const {
    _service: dataverseService,
    deleteButtonVisibility,
    newButtonVisibility,
    defaultPageSize,
    fetchXml,
    setIsLoading,
    setError,
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
        const filteredColumns = await getSortedColumns(fetchXml, dataverseService);
        setColumns(filteredColumns);
        displayName.current = await dataverseService.getEntityDisplayName(entityName);
      }
      catch (err: any) {
        setError(err);
      }
    })();
  }, [fetchXml, deleteButtonVisibility, newButtonVisibility, pageSize]);

  React.useEffect(() => {
    deleteBtnClassName.current = 'disableButton';
    (async () => {
      setError();
      setIsLoading(true);
      if (isDialogAccepted) return;

      try {
        totalRecordsCount.current = await dataverseService.getRecordsCount(fetchXml ?? '');
        const records: Entity[] = await getItems(
          fetchXml,
          pageSize,
          currentPage,
          totalRecordsCount.current,
          dataverseService);

        calculateFilteredRecordsData(
          totalRecordsCount.current,
          records,
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
        setItems(records);
      }
      catch (err: any) {
        setError(err);
      }

      setIsLoading(false);
    })();
  }, [fetchXml, pageSize, currentPage, isDialogAccepted]);

  return <>
    <Stack horizontal horizontalAlign="end" className={dataSetStyles.buttons}>
      <CommandBar
        _service={dataverseService}
        entityName={entityName}
        selectedRecordIds={selectedRecordIds}
        displayName={displayName.current}
        setDialogAccepted={setDialogAccepted}
        className={deleteBtnClassName.current}
        deleteButtonVisibility={deleteButtonVisibility}
        newButtonVisibility={newButtonVisibility}
      />
    </Stack>

    <List _service={dataverseService}
      entityName={entityName}
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
  </>
  ;
};
