import * as React from 'react';
import { IAppWrapperProps } from '../@types/types';
import { FetchToSubgrid } from './FetchToSubgrid';
import { InfoMessage } from './InfoMessage';
import { Loader } from './Loader';

export const AppWrapper: React.FC<IAppWrapperProps> = props => {
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState<Error | undefined>(undefined);

  React.useEffect(() => {
    setError(props.error);
  }, [props.error]);

  return <>
    { error && <InfoMessage error={error} dataverseService={props._service} /> }
    <div className='FetchToSubgridControl' style={{ display: error ? 'none' : 'grid' }} >
      {isLoading && <Loader />}
      <FetchToSubgrid
        {...props}
        setIsLoading={setIsLoading}
        setError={setError}
      />
    </div>
  </>;
};
