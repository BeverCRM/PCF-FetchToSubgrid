import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';

export const Loader: React.FC = () =>
  <div className='fetchLoading'>
    <Spinner size={SpinnerSize.large} />
    <p className='loadingText'>  Loading ...</p>
  </div>;
