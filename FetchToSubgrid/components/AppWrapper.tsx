import * as React from 'react';
import { FetchSubgrid } from './FetchSubgrid';
import { InfoMessage } from './InfoMessage';
import { Loader } from './Loader';

export interface IAppWrapperProps {
  fetchXml: string | null;
  defaultPageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
}

export const AppWrapper: React.FC<IAppWrapperProps> = props => {
  const [isLoading, setIsLoading] = React.useState(false);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  return <>
    { errorMessage && <InfoMessage message={errorMessage }/> }
    { isLoading && <Loader /> }
    <FetchSubgrid
      {...props}
      isVisible={!isLoading && !errorMessage}
      setIsLoading={setIsLoading}
      setErrorMessage={setErrorMessage}
    />
  </>;
};
