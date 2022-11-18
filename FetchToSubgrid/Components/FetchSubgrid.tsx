import * as React from 'react';
import { DetailsList, DetailsListLayoutMode,
  Spinner, SpinnerSize } from '@fluentui/react';
import FetchService from '../Services/FetchService';
import { useState, useEffect, useRef } from 'react';

export interface IFetchShbgridProps {
  inputValue: string | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchShbgridProps> = props => {
  const [ isLoading, setIsLoading ] = useState<boolean>(true);
  const hasData = useRef<boolean>(false);
  const entities = useRef([]);

  useEffect(() => {
    setIsLoading(true);
    FetchService.getFetchData().then(
      (records: any) => {
        entities.current = records;
        hasData.current = true;
        setIsLoading(false);
      },
      (error: any) => {
        entities.current = [];
        hasData.current = false;
        console.error(error);
        setIsLoading(false);
      });
  }, [props.inputValue]);

  if (isLoading) {
    return (
      <div className='fetchSubgridControl'>
        <div className='fetchLoading'>
          <Spinner size={SpinnerSize.large} />
          <p className='loadingText'>  Loading ...</p>
        </div>
      </div>);
  }

  if (!hasData.current) {
    return <div className='fetchSubgridControl'>
      <div className='infoMessage'>
        <h1 className='infoMessageText'>No data available</h1>
      </div>
    </div>;
  }

  return (
    <div className='fetchSubgridControl'>
      <DetailsList
        items={entities.current}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        styles={{ contentWrapper: { maxHeight: 200 } }}
      >
      </DetailsList>
    </div>
  );
};
