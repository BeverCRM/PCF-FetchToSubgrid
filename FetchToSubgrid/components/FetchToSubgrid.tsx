import * as React from 'react';
import { IColumn, IObjectWithKey, Stack } from '@fluentui/react';
import { Entity } from '../@types/types';
import { getEntityNameFromFetchXml, isAggregate } from '../utilities/fetchXmlUtils';
import { dataSetStyles } from '../styles/comandBarStyles';
import { CommandBar } from './ComandBar';
import { List } from './List';
import { getItems } from '../utilities/d365Utils';
import { Footer } from './Footer';
import { useSelection } from '../hooks/useSelection';
import { IDataverseService } from '../services/dataverseService';
import { getSortedColumns, createLinkableItems, hashCode } from '../utilities/utils';

export interface IFetchToSubgridProps {
  fetchXml: string | null;
  _service: IDataverseService;
  pageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
  allocatedWidth: number;
  error?: Error;
  setIsLoading: (isLoading: boolean) => void;
  setError: (error?: Error | undefined) => void;
}

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
  const displayName = React.useRef('');
  const isButtonActive = React.useRef(false);
  const totalRecordsCount = React.useRef(0);
  const selectedItemsCount = React.useRef(0);
  const fetchXmlOldValue = React.useRef<string | null>(null);

  const entityName = getEntityNameFromFetchXml(fetchXml ?? '');
  const firstItemIndex = (currentPage - 1) * pageSize + 1;
  const lastItemIndex = (currentPage - 1) * pageSize + items.length;
  const nextButtonDisabled = Math.ceil(totalRecordsCount.current / pageSize) <= currentPage;

  const { selection, selectedRecordIds } = useSelection((currentSelection: IObjectWithKey[]) => {
    selectedItemsCount.current = currentSelection.length;
    isButtonActive.current = currentSelection.length > 0 && !isAggregate(fetchXml);
  });

  React.useEffect(() => {
    const fetchColumns = async () => {
      try {
        const filteredColumns = await getSortedColumns(fetchXml, allocatedWidth, dataverseService);
        setColumns(filteredColumns);
        listInputsHashCode.current = hashCode(`${allocatedWidth}${fetchXml}`);
        displayName.current = await dataverseService.getEntityDisplayName(entityName);
        if (fetchXml !== fetchXmlOldValue.current) {
          fetchXmlOldValue.current = fetchXml;
          setCurrentPage(1);
        }
      }
      catch (error: any) {
        setError(error);
      }
    };

    fetchColumns();
  }, [fetchXml, allocatedWidth]);

  React.useEffect(() => {
    const fetchItems = async () => {
      isButtonActive.current = false;
      setIsLoading(true);
      if (isDialogAccepted) return;

      try {
        if (fetchXml !== fetchXmlOldValue.current) {
          totalRecordsCount.current = await dataverseService.getRecordsCount(fetchXml ?? '');
        }

        const records: Entity[] = await getItems(
          fetchXml,
          pageSize,
          currentPage,
          totalRecordsCount.current,
          dataverseService);

        const linkableItems = createLinkableItems(records, recordIds.current, dataverseService);

        setItems(linkableItems);
      }
      catch (error: any) {
        setError(error);
      }
      setIsLoading(false);
    };

    fetchItems();
  }, [currentPage, isDialogAccepted, fetchXml, pageSize]);

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
      firstItemIndex={firstItemIndex}
      lastItemIndex={lastItemIndex}
      selectedItemsCount={selectedItemsCount.current}
      totalRecordsCount={totalRecordsCount.current}
      currentPage={currentPage}
      setCurrentPage={setCurrentPage}
      nextButtonDisable={nextButtonDisabled}
      movePreviousIsDisabled={currentPage <= 1}
    />
  </>;
};
