import * as React from 'react';
import { DetailsList, DetailsListLayoutMode,
  Spinner, SpinnerSize } from '@fluentui/react';
import FetchService from '../Services/CrmService';
import { useState, useEffect } from 'react';
import LinkableItems from './LinkableItems';

export interface IFetchSubgridProps {
  fetchXml: string | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchSubgridProps> = props => {
  const [ isLoading, setIsLoading ] = useState(true);
  const { fetchXml } = props;
  const [items, setItems] = useState([]);
  const [columns, setColumns] = useState([]);

  useEffect(() => {
    (async () => {
      setIsLoading(true);

      try {
        const columns: any = await FetchService.getColumns(fetchXml);
        setColumns(columns);

        const items: any = await LinkableItems.getLinkableItems(fetchXml);
        setItems(items);
      }
      catch (err) {
        setColumns([]);
        setItems([]);
      }

      setIsLoading(false);
    })();
  }, [props]);

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
        styles={{ contentWrapper: { maxHeight: 200 } }}
      >
      </DetailsList>
    </div>
  );
};
