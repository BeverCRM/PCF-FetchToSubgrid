import * as React from 'react';
import { DetailsList, DetailsListLayoutMode,
  Spinner, SpinnerSize } from '@fluentui/react';
import FetchService from '../Services/CrmService';
import { useState, useEffect } from 'react';

export interface IFetchSubgridProps {
  defaultFetchXml: string | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchSubgridProps> = props => {
  const [ isLoading, setIsLoading ] = useState<boolean>(true);
  const { defaultFetchXml } = props;
  const [items, setItems] = useState([]);
  const [columns, setColumns] = useState([]);

  useEffect(() => {
    setIsLoading(true);

    FetchService.getColumns(defaultFetchXml).then(
      (columns: any) => {
        setColumns(columns);
      },
      (error: any) => {
        setColumns([]);
        console.error(error);
      });

    FetchService.getItems(defaultFetchXml).then(
      (items: any) => {
        setItems(items);
        setIsLoading(false);
      },
      (error: any) => {
        setItems([]);
        setIsLoading(false);
        console.error(error);
      });

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
