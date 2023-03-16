import * as React from 'react';
import { IColumn, IObjectWithKey, Stack } from '@fluentui/react';
import { Entity, IFetchToSubgridProps } from '../@types/types';
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
  const listInputsHashCode = React.useRef(-1);

  const recordIds = React.useRef<string[]>([]);
  const nextButtonDisabled = React.useRef(true);
  const displayName = React.useRef('');
  const isButtonActive = React.useRef(false);
  const totalRecordsCount = React.useRef(0);
  const selectedItemsCount = React.useRef(0);
  const firstItemIndex = React.useRef(0);
  const lastItemIndex = React.useRef(0);
  const fetchXmlOldValue = React.useRef(fetchXml);

  const entityName = getEntityNameFromFetchXml(fetchXml ?? '');

  const { selection, selectedRecordIds } = useSelection((currentSelection: IObjectWithKey[]) => {
    selectedItemsCount.current = currentSelection.length;
    isButtonActive.current = currentSelection.length > 0 && !isAggregate(fetchXml);
  });

  React.useEffect(() => {
    (async () => {
      try {
        const filteredColumns = await getSortedColumns(fetchXml, allocatedWidth, dataverseService);
        setColumns(filteredColumns);
        listInputsHashCode.current = `${allocatedWidth}${fetchXml}`.hashCode();
        if (fetchXml !== fetchXmlOldValue.current) {
          setCurrentPage(1);
          displayName.current = await dataverseService.getEntityDisplayName(entityName);
          fetchXmlOldValue.current = fetchXml;
        }
      }
      catch (error: any) {
        setError(error);
      }
    })();
  }, [fetchXml, allocatedWidth]);

  React.useEffect(() => {
    (async () => {
      isButtonActive.current = false;
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
      catch (error: any) {
        setError(error);
      }
      setIsLoading(false);
    })();
  },
  [
    currentPage,
    isDialogAccepted,
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
      pageSize={pageSize}
      forceReRender={listInputsHashCode.current}
      firstItemIndex={firstItemIndex}
      lastItemIndex={lastItemIndex}
      selectedItemsCount={selectedItemsCount}
      totalRecordsCount={totalRecordsCount.current}
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
      selectedItemsCount={selectedItemsCount.current}
      totalRecordsCount={totalRecordsCount.current}
      currentPage={currentPage}
      setCurrentPage={setCurrentPage}
      nextButtonDisable={nextButtonDisabled.current}
      movePreviousIsDisabled={currentPage <= 1}
    />
  </>;
};
