import * as React from 'react';
import { IInfoMessageProps } from '../utilities/types';

export const InfoMessage: React.FC<IInfoMessageProps> = ({ error, dataverseService }) =>
  <div className='FetchToSubgridControl'>
    <div className='infoMessage'>
      <h1
        className='infoMessageText' onClick={ () =>
          dataverseService.showNotificationPopup(error)}
      >
        An error has occurred!
      </h1>
    </div>
  </div>;
