import * as React from 'react';
import { DetailsList } from '@fluentui/react';
import FetchService from '../Services/FetchService';

export interface IFetchShbgridProps {
  inputValue: string | null;
}

export const FetchSubgrid: React.FunctionComponent<IFetchShbgridProps> = props => {

  const [ entities, setEntities ] = React.useState([]);
  const [ hasData, setHasData ] = React.useState(false);

  React.useEffect(() => {
    FetchService.getFetchData().then(
      (records: any) => {
        setHasData(true);
        setEntities(records);
      },
      (error: any) => {
        console.error(error);
        setHasData(false);
      });
  }, [props.inputValue]);

  if (!hasData) return <span>Error</span>;

  return (
    <DetailsList
      items={entities}
      styles={{ contentWrapper: { minHeight: 150 } }}
    >
    </DetailsList>
  );
};
