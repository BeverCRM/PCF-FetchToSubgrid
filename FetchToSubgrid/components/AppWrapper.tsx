import * as React from 'react';
import { IAppWrapperProps } from '../utilities/types';
import { FetchSubgrid } from './FetchSubgrid';
import { Loader } from './Loader';

export const AppWrapper: React.FC<IAppWrapperProps> = props => {
  const [isLoading, setIsLoading] = React.useState(false);

  return <div className='FetchSubgridControl'>
    { isLoading && <Loader /> }
    <FetchSubgrid
      {...props}
      setIsLoading={setIsLoading}
    />
  </div>;
};
