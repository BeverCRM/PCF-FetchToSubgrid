import * as React from 'react';

interface IInfoMessageProps {
  message: string;
}

export const InfoMessage: React.FC<IInfoMessageProps> = ({ message }) =>
  <div className='fetchSubgridControl'>
    <div className='infoMessage'>
      <h1 className='infoMessageText'>
        {message ?? 'No data available'}
      </h1>
    </div>
  </div>;
