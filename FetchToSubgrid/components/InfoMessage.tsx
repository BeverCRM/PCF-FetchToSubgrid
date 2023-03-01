import * as React from 'react';
import { IInfoMessageProps } from '../utilities/types';

export const InfoMessage: React.FC<IInfoMessageProps> = ({ message }) =>
  <div className='FetchToSubgridControl'>
    <div className='infoMessage'>
      <h1 className='infoMessageText'>
        {message || 'No data available'}
      </h1>
    </div>
  </div>;
