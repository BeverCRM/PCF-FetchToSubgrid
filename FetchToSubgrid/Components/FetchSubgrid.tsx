import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IDetailsFooterProps,
  IDetailsListProps, Spinner, SpinnerSize } from '@fluentui/react';
import FetchService from '../Services/CrmService';
import { useState, useEffect, useCallback } from 'react';
import LinkableItems from './LinkableItems';
import CrmService from '../Services/CrmService';
import { GridFooter } from './Footer';
import utilities from '../utilities';

export interface IFetchSubgridProps {
  fetchXml: string | null;
  numberOfRows: number | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchSubgridProps> = props => {
  const [isLoading, setIsLoading] = useState(true);
  const { numberOfRows, fetchXml } = props;
  const [items, setItems] = useState<any>([]);
  const [columns, setColumns] = useState([]);
  const [currentPage, setCurrentPage] = useState(1);
  let nextButtonDisable = true;
  let recordsPerPage: number = CrmService.getPagingLimit();

  if (numberOfRows) {
    recordsPerPage = numberOfRows;
  }

  React.useEffect(() => {
    (async () => {
      const recordsCount = await CrmService.getRecordsCount(fetchXml ?? '');
      if (Math.ceil(recordsCount / recordsPerPage) <= currentPage) {
        nextButtonDisable = true;
      }
      else {
        nextButtonDisable = false;
      }
    })();
  }, [fetchXml]);

  const onRenderDetailsFooter: IDetailsListProps['onRenderDetailsFooter'] =
    (props: IDetailsFooterProps | undefined) => {
      const isMovePrevious = !(currentPage > 1);

      if (props) {
        return <GridFooter
          currentPage={currentPage}
          setCurrentPage={setCurrentPage}
          nextButtonDisable={nextButtonDisable}
          isMovePrevious={isMovePrevious}
        ></GridFooter>;
      }
      return null;
    };

  const onItemInvoked = useCallback((item : any) : void => {
    for (let i = 0; i < items.length; i++) {
      if (item.key === items[i].key) {
        CrmService.openRecord(item.entityName, item.key);
      }
    }
  }, [items]);

  useEffect(() => {
    (async () => {
      setIsLoading(true);
      try {
        const columns: any = await FetchService.getColumns(fetchXml);
        setColumns(columns);
        const items: any = await utilities.getItems(fetchXml, recordsPerPage, currentPage);

        items.forEach((item: any) => {
          Object.keys(item).forEach(key => {
            const value = item[key];
            item[key] = value.linkable ? LinkableItems.makeItemsLinkable(value) : value.displayName;
          });
        });

        setItems(items);
      }
      catch (err) {
        setColumns([]);
        setItems([]);
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
        onItemInvoked= {onItemInvoked}
        onRenderDetailsFooter={onRenderDetailsFooter}
      >
      </DetailsList>
    </div>
  );
};
