import * as React from 'react';
import { IAppWrapperProps } from '../@types/types';
import { FetchToSubgrid } from './FetchToSubgrid';
import { InfoMessage } from './InfoMessage';
import { Loader } from './Loader';

export const AppWrapper: React.FC<IAppWrapperProps> = props => {
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState<Error | undefined>(undefined);

  React.useEffect(() => {
    if (props.error === error) setError(undefined);
    else setError(props.error);
  }, [props.fetchXml, props.error]);

  return (
    <div className='FetchToSubgridControl'>
      { error && <InfoMessage error={error} dataverseService={props._service} /> }
      { !error && isLoading && <Loader />}
      { !error &&
        <FetchToSubgrid
          {...props}
          setIsLoading={setIsLoading}
          setError={setError}
        />
      }
    </div>
  );
};
