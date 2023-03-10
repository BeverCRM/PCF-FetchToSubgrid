import * as React from 'react';
import { IColumn, IObjectWithKey, Stack } from '@fluentui/react';
import { Entity, IFetchToSubgridProps } from '../utilities/types';
import { getSortedColumns, calculateFilteredRecordsData } from '../utilities/utils';
import { getEntityNameFromFetchXml, isAggregate } from '../utilities/fetchXmlUtils';
import { dataSetStyles } from '../styles/comandBarStyles';
import { LinkableItem } from './LinkableItems';
import { CommandBar } from './ComandBar';
import { List } from './List';
import { getItems } from '../utilities/d365Utils';
import { Footer } from './Footer';
import { useSelection } from '../hooks/useSelection';

export const FetchToSubgrid: React.FC<IFetchToSubgridProps> = props => {
  const {
    _service: dataverseService,
    deleteButtonVisibility,
    newButtonVisibility,
    pageSize,
    allocatedWidth,
    fetchXml,
    setIsLoading,
    setError,
  } = props;

  const [items, setItems] = React.useState<Entity[]>([]);
  const [currentPage, setCurrentPage] = React.useState(1);
  const [isDialogAccepted, setDialogAccepted] = React.useState(false);
  const [columns, setColumns] = React.useState<IColumn[]>([]);
  const [allocatedWidthKey, setAllocatedWidthKey] = React.useState(allocatedWidth);

  const recordIds = React.useRef<string[]>([]);
  const nextButtonDisabled = React.useRef(true);
  const displayName = React.useRef('');
  const isButtonActive = React.useRef(false);
  const totalRecordsCount = React.useRef(0);
  const selectedItemsCount = React.useRef(0);
  const firstItemIndex = React.useRef(0);
  const lastItemIndex = React.useRef(0);

  const entityName = getEntityNameFromFetchXml(fetchXml ?? '');

  const updatedFetchXml = React.useRef(fetchXml);
  React.useEffect(() => { updatedFetchXml.current = fetchXml; }, [fetchXml]);

  const { selection, selectedRecordIds } = useSelection((currentSelection: IObjectWithKey[]) => {
    selectedItemsCount.current = currentSelection.length;
    isButtonActive.current = currentSelection.length > 0 &&
      !isAggregate(updatedFetchXml.current);
  });

  React.useEffect(() => {
    (async () => {
      try {
        setCurrentPage(1);
        const filteredColumns = await getSortedColumns(fetchXml, allocatedWidth, dataverseService);
        setColumns(filteredColumns);
        setAllocatedWidthKey(allocatedWidth);
        displayName.current = await dataverseService.getEntityDisplayName(entityName);
      }
      catch {
        const error = await dataverseService.getCurrentPageRecords(fetchXml);
        setError(error);
      }
    })();
  }, [fetchXml, allocatedWidth]);

  React.useEffect(() => {
    (async () => {
      isButtonActive.current = false;
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
        const error = await dataverseService.getCurrentPageRecords(fetchXml);
        setError(error);
      }

      setIsLoading(false);
    })();
  },
  [
    currentPage,
    isDialogAccepted,
    // deleteButtonVisibility,
    // newButtonVisibility,
    fetchXml,
  ]);

  return <>
    <Stack horizontal horizontalAlign="end" className={dataSetStyles.buttons}>
      <CommandBar
        _service={dataverseService}
        entityName={entityName}
        selectedRecordIds={selectedRecordIds}
        displayName={displayName.current}
        setDialogAccepted={setDialogAccepted}
        isButtonActive={isButtonActive.current}
        deleteButtonVisibility={deleteButtonVisibility}
        newButtonVisibility={newButtonVisibility}
      />
    </Stack>

    <List _service={dataverseService}
      entityName={entityName}
      isButtonActive={isButtonActive}
      pageSize={pageSize}
      allocatedWidthKey={allocatedWidthKey}
      firstItemIndex={firstItemIndex}
      lastItemIndex={lastItemIndex}
      selectedItemsCount={selectedItemsCount}
      totalRecordsCount={totalRecordsCount}
      nextButtonDisabled={nextButtonDisabled}
      fetchXml={fetchXml}
      recordIds={recordIds}
      selection={selection}
      currentPage={currentPage}
      setCurrentPage={setCurrentPage}
      columns={columns}
      setColumns={setColumns}
      items={items}
      setItems={setItems}
    />

    <Footer
      firstItemIndex={firstItemIndex.current}
      lastItemIndex={lastItemIndex.current}
      selectedItems={selectedItemsCount.current}
      totalRecordsCount={totalRecordsCount.current}
      currentPage={currentPage}
      setCurrentPage={setCurrentPage}
      nextButtonDisable={nextButtonDisabled.current}
      movePreviousIsDisabled={currentPage <= 1}
    />
  </>;
};
