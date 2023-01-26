import * as React from 'react';

interface IInfoMessageProps {
  fetchXml: string | null;
}

export const InfoMessage: React.FC<IInfoMessageProps> = ({ fetchXml }) =>
  <div className='fetchSubgridControl'>
    <div className='infoMessage'>
      { fetchXml
        ? <h1 className='infoMessageText'>No data available</h1>
        : <h1 className='infoMessageText'>
          Default FetchXml is not specified and the field is blank.
        </h1>
      }
    </div>
  </div>;
