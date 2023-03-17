import * as React from 'react';
import { IAppWrapperProps, IFetchToSubgridProps } from '../@types/types';
import { parseRawInput } from '../utilities/utils';
import { FetchToSubgrid } from './FetchToSubgrid';
import { InfoMessage } from './InfoMessage';
import { Loader } from './Loader';

export const AppWrapper: React.FC<IAppWrapperProps> = props => {
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState<Error | undefined>(undefined);

  const fetchToSubgridProps: IFetchToSubgridProps = parseRawInput(props, setIsLoading, setError);
  if (fetchToSubgridProps.error) setError(error);

  React.useEffect(() => {
    setError(undefined);
  }, [props.fetchXmlOrJson]);

  return (
    <div className='FetchToSubgridControl'>
      { error && <InfoMessage error={error} dataverseService={props._service} /> }
      { !error && isLoading && <Loader />}
      { !error &&
        <FetchToSubgrid
          {...fetchToSubgridProps}
          setIsLoading={setIsLoading}
          setError={setError}
        />
      }
    </div>
  );
};
