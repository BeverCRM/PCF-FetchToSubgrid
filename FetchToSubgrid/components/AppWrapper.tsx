import * as React from 'react';
import { IService } from '../@types/types';
import { IDataverseService } from '../services/dataverseService';
import { parseRawInput } from '../utilities/utils';
import { FetchToSubgrid, IFetchToSubgridProps } from './FetchToSubgrid';
import { ErrorMessage } from './ErrorMessage';
import { Loader } from './Loader';

export interface IAppWrapperProps extends IService<IDataverseService> {
  fetchXmlOrJson: string | null;
  allocatedWidth: number;
  default: {
    fetchXml: string | null;
    pageSize: number;
    deleteButtonVisibility: boolean;
    newButtonVisibility: boolean;
  }
}

export const AppWrapper: React.FC<IAppWrapperProps> = props => {
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState<Error | undefined>(undefined);

  const fetchToSubgridProps: IFetchToSubgridProps = parseRawInput(props);

  React.useEffect(() => {
    if (fetchToSubgridProps.error) {
      setError(fetchToSubgridProps.error);
    }
    else {
      setError(undefined);
    }
  }, [props.fetchXmlOrJson]);

  return (
    <div className='FetchToSubgridControl'>
      { error && <ErrorMessage error={error} dataverseService={props._service} /> }
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
