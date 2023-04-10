import * as React from 'react';
import { IColumn, Stack } from '@fluentui/react';
import { Entity, IItemsData } from '../@types/types';
import { dataSetStyles } from '../styles/comandBarStyles';
import { CommandBar } from './ComandBar';
import { List } from './List';
import { Footer } from './Footer';
import { useSelection } from '../hooks/useSelection';
import { IDataverseService } from '../services/dataverseService';
import { getSortedColumns, hashCode, getLinkableItems } from '../utilities/utils';
import {
  addOrderToFetch,
  getEntityNameFromFetchXml,
  isAggregate } from '../utilities/fetchXmlUtils';

export interface IFetchToSubgridProps {
  fetchXml: string | null;
  _service: IDataverseService;
  pageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
  allocatedWidth: number;
  error?: Error;
  setIsLoading?: (isLoading: boolean) => void;
  setError?: (error?: Error | undefined) => void;
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
  const totalRecordsCount = React.useRef(0);
  const [sortingData, setSortingData] = React.useState({ fieldName: '', column: undefined });

  const entityName = getEntityNameFromFetchXml(fetchXml ?? '');

  const { selection, selectedRecordIds } = useSelection();
  let isButtonActive = selectedRecordIds.length > 0 && !isAggregate(fetchXml);

  React.useLayoutEffect(() => {
    if (allocatedWidth === -1) return;

    listInputsHashCode.current = hashCode(`${allocatedWidth}${fetchXml}`);
  }, [allocatedWidth]);

  React.useMemo(async () => {
    totalRecordsCount.current = await dataverseService.getRecordsCount(fetchXml ?? '');
  }, [fetchXml, isDialogAccepted]);

  React.useEffect(() => setCurrentPage(1), [pageSize, fetchXml]);

  React.useEffect(() => {
    const fetchColumns = async () => {
      try {
        if (allocatedWidth === -1) return;

        const filteredColumns = await getSortedColumns(fetchXml, allocatedWidth, dataverseService);
        setColumns(filteredColumns);
      }
      catch (error: any) {
        setError!(error);
      }
    };

    fetchColumns();
  }, [fetchXml, allocatedWidth]);

  React.useEffect(() => {
    const fetchItems = async () => {
      isButtonActive = false;
      setIsLoading!(true);
      if (isDialogAccepted) return;

      try {
        const newFetchXml = addOrderToFetch(fetchXml, sortingData);

        const data: IItemsData = {
          fetchXml: sortingData.column === undefined ? fetchXml : newFetchXml,
          pageSize,
          currentPage,
        };

        const linkableItems: Entity[] = await getLinkableItems(data, dataverseService);

        setItems(linkableItems);
      }
      catch (error: any) {
        setError!(error);
      }
      setIsLoading!(false);
    };

    fetchItems();
  }, [currentPage, isDialogAccepted, fetchXml, pageSize, sortingData]);

  return <>
    <Stack horizontal horizontalAlign="end" className={dataSetStyles.buttons}>
      <CommandBar
        _service={dataverseService}
        entityName={entityName}
        selectedRecordIds={selectedRecordIds}
        setDialogAccepted={setDialogAccepted}
        isButtonActive={isButtonActive}
        deleteButtonVisibility={deleteButtonVisibility}
        newButtonVisibility={newButtonVisibility}
      />
    </Stack>

    <List
      _service={dataverseService}
      entityName={entityName}
      forceReRender={listInputsHashCode.current}
      fetchXml={fetchXml}
      selection={selection}
      columns={columns}
      items={items}
      setSortingData={setSortingData}
    />

    <Footer
      pageSize={pageSize}
      selectedItemsCount={selectedRecordIds.length}
      totalRecordsCount={totalRecordsCount.current}
      currentPage={currentPage}
      setCurrentPage={setCurrentPage}
    />
  </>;
};
