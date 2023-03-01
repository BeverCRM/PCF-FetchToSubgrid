import * as React from 'react';
import { IAppWrapperProps } from '../utilities/types';
import { FetchToSubgrid } from './FetchToSubgrid';
import { Loader } from './Loader';

export const AppWrapper: React.FC<IAppWrapperProps> = props => {
  const [isLoading, setIsLoading] = React.useState(false);

  return <div className='FetchToSubgridControl'>
    { isLoading && <Loader /> }
    <FetchToSubgrid
      {...props}
      setIsLoading={setIsLoading}
    />
  </div>;
};
