import { Spinner, SpinnerSize } from '@fluentui/react';
import * as React from 'react';

export const Loader: React.FC = () =>
  <div className='FetchSubgridControl'>
    <div className='fetchLoading'>
      <Spinner size={SpinnerSize.large} />
      <p className='loadingText'>  Loading ...</p>
    </div>
  </div>;
