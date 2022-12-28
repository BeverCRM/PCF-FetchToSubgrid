import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsFooterProps,
  IDetailsListProps,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { LinkableItem } from './LinkableItems';
import { getPagingLimit, getColumns, getRecordsCount, openRecord } from '../Services/CrmService';
import { Footer } from './Footer';
import { getCountInFetchXml, getEntityName, getItems } from '../Utilities/utilities';

export interface IFetchSubgridProps {
  fetchXml: string | null;
  numberOfRows: number | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchSubgridProps> = props => {
  const { numberOfRows, fetchXml } = props;

  const [isLoading, setIsLoading] = React.useState(false);
  const [ columns, setColumns] = React.useState<IColumn[]>([]);
  const [currentPage, setCurrentPage] = React.useState(1);
  const [items, setItems] = React.useState<ComponentFramework.WebApi.Entity[]>([]);

  const recordIds = React.useRef<string[]>([]);

  const nextButtonDisable = React.useRef(true);
  let recordsPerPage: number = getPagingLimit();
  const countOfRecordsInFetch = getCountInFetchXml(fetchXml);

  if (countOfRecordsInFetch) {
    recordsPerPage = countOfRecordsInFetch;
  }
  else if (numberOfRows) {
    recordsPerPage = numberOfRows;
  }

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = React.useCallback(
    (props: IDetailsFooterProps | undefined) => {
      const isMovePrevious = !(currentPage > 1);
      if (props) {
        return <Footer
          currentPage={currentPage}
          setCurrentPage={setCurrentPage}
          nextButtonDisable={nextButtonDisable.current}
          isMovePrevious={isMovePrevious}
        >
        </Footer>;
      }
      return null;
    },
    [currentPage, nextButtonDisable],
  );

  const onItemInvoked = React.useCallback((
    record: ComponentFramework.WebApi.Entity,
    index?: number | undefined) : void => {
    const entityName = getEntityName(fetchXml ?? '');

    if (index !== undefined) {
      openRecord(entityName, recordIds.current[index]);
    }
  }, [items]);

  React.useEffect(() => {
    (async () => {
      setCurrentPage(1);
      try {
        const columns = await getColumns(fetchXml);
        setColumns(columns);
      }
      catch {
        setColumns([]);
      }
    })();
  }, [fetchXml]);

  React.useEffect(() => {
    (async () => {
      setIsLoading(true);
      try {
        const recordsCount = await getRecordsCount(fetchXml ?? '');
        if (Math.ceil(recordsCount / recordsPerPage) > currentPage) {
          nextButtonDisable.current = false;
        }
        else {
          nextButtonDisable.current = true;
        }

        const records: ComponentFramework.WebApi.Entity[] = await getItems(
          fetchXml,
          recordsPerPage,
          currentPage);

        records.forEach(record => {
          recordIds.current.push(record.id);

          Object.keys(record).forEach(key => {
            const value: any = record[key];
            record[key] = value.linkable ? <LinkableItem item = {value} /> : value.displayName;
          });
        });

        setItems(records);
      }
      catch (err) {
        console.log('Error', err);
      }
      setIsLoading(false);
    })();

  }, [props, currentPage]);

  if (isLoading) {
    return (
      <div className='fetchSubgridControl'>
        <div className='fetchLoading'>
          <Spinner size={SpinnerSize.large} />
          <p className='loadingText'>  Loading ...</p>
        </div>
      </div>);
  }

  if (columns.length === 0) {
    return <div className='fetchSubgridControl'>
      <div className='infoMessage'>
        <h1 className='infoMessageText'>No data available</h1>
      </div>
    </div>;
  }

  return (
    <div className='fetchSubgridControl'>
      <DetailsList
        columns={columns}
        items={items}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        onItemInvoked={onItemInvoked}
        onRenderDetailsFooter={onRenderDetailsFooter}
      >
      </DetailsList>
    </div>
  );
};
