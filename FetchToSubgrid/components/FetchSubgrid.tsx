import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsFooterProps,
  IDetailsListProps,
} from '@fluentui/react';
import { LinkableItem } from './LinkableItems';
import { getPagingLimit, getColumns, getRecordsCount, openRecord } from '../Services/CrmService';
import { Footer } from './Footer';
import { getCountInFetchXml, getEntityName, getItems, isAggregate } from '../Utilities/utilities';
import { Loader } from './Loader';
import { InfoMessage } from './InfoMessage';

export interface IFetchSubgridProps {
  fetchXml: string | null;
  numberOfRows: number | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchSubgridProps> = props => {
  const { numberOfRows, fetchXml } = props;

  const [isLoading, setIsLoading] = React.useState(false);
  const [columns, setColumns] = React.useState<IColumn[]>([]);
  const [currentPage, setCurrentPage] = React.useState(1);
  const [items, setItems] = React.useState<ComponentFramework.WebApi.Entity[]>([]);

  const recordIds = React.useRef<string[]>([]);
  const nextButtonDisable = React.useRef(true);
  let recordsPerPage: number = getPagingLimit();
  const countOfRecordsInFetch: number = getCountInFetchXml(fetchXml);

  recordsPerPage = countOfRecordsInFetch || numberOfRows || recordsPerPage;

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = React.useCallback(
    (props: IDetailsFooterProps | undefined) => {
      const isMovePrevious = !(currentPage > 1);
      if (props) {
        return (
          <Footer
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
  }, [items]);

  React.useEffect(() => {
    (async () => {
      setCurrentPage(1);
      try {
        const columns: IColumn[] = await getColumns(fetchXml);
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
        const recordsCount: number = await getRecordsCount(fetchXml ?? '');
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
            if (key !== 'id') {
              const value: any = record[key];
              record[key] = value.linkable ? <LinkableItem item = {value} /> : value.displayName;
            }
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
    return <Loader/>;
  }

  if (columns.length === 0) {
    return <InfoMessage/>;
  }

  return (
    <div className='fetchSubgridControl'>
      <DetailsList
        columns={columns}
        items={items}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        onItemInvoked={onItemInvoked}
        onRenderDetailsFooter={onRenderDetailsFooter}
      />
    </div>
  );
};
