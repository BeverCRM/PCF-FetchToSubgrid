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
import { useState, useEffect, useCallback } from 'react';
import { LinkableItem } from './LinkableItems';
import { getPagingLimit, getColumns, getRecordsCount, openRecord } from '../Services/CrmService';
import { GridFooter } from './Footer';
import { getEntityName, getItems } from '../Utilities/utilities';

export interface IFetchSubgridProps {
  fetchXml: string | null;
  numberOfRows: number | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchSubgridProps> = props => {
  const [isLoading, setIsLoading] = useState(false);
  const { numberOfRows, fetchXml } = props;
  const items = React.useRef<ComponentFramework.WebApi.Entity[]>([]);
  const [ columns, setColumns] = useState<IColumn[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const recordIds = React.useRef<string[]>([]);

  let nextButtonDisable = true;
  let recordsPerPage = getPagingLimit();

  if (numberOfRows) {
    recordsPerPage = numberOfRows;
  }

  React.useEffect(() => {
    (async () => {
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
      const recordsCount = await getRecordsCount(fetchXml ?? '');

      Math.ceil(recordsCount / recordsPerPage) > currentPage ? nextButtonDisable = false
        : nextButtonDisable = true;

    })();
  }, [currentPage]);

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] = useCallback(
    (props: IDetailsFooterProps | undefined) => {
      const isMovePrevious = !(currentPage > 1);

      if (props) {
        return <GridFooter
          currentPage={currentPage}
          setCurrentPage={setCurrentPage}
          nextButtonDisable={nextButtonDisable}
          isMovePrevious={isMovePrevious}
        >
        </GridFooter>;
      }
      return null;
    },
    [currentPage, nextButtonDisable, setCurrentPage],
  );

  const onItemInvoked = useCallback((
    item: ComponentFramework.WebApi.Entity,
    index?: number | undefined) : void => {
    const entityName = getEntityName(fetchXml ?? '');

    if (index !== undefined) {
      openRecord(entityName, recordIds.current[index]);
    }
  }, [items]);

  useEffect(() => {
    (async () => {
      setIsLoading(true);
      try {
        items.current = await getItems(fetchXml, recordsPerPage, currentPage);
        items.current.forEach(item => {
          recordIds.current.push(item.id);
          Object.keys(item).forEach(key => {
            const value: ComponentFramework.WebApi.Entity = item[key];
            item[key] = value.linkable ? <LinkableItem item = {value} /> : value.displayName;
          });
        });
        setIsLoading(false);
      }
      catch {
        setIsLoading(false);
      }
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
        items={items.current}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        onItemInvoked={onItemInvoked}
        onRenderDetailsFooter={onRenderDetailsFooter}
      >
      </DetailsList>
    </div>
  );
};
