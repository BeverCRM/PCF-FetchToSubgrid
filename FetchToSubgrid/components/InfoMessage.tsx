import * as React from 'react';

interface IInfoMessageProps {
  fetchXml: string | null;
}

export const InfoMessage: React.FC<IInfoMessageProps> = ({ fetchXml }) =>
  <div className='fetchSubgridControl'>
    <div className='infoMessage'>
      <h1 className='infoMessageText'>
        {fetchXml
          ? 'No data available'
          : 'Default FetchXml is not specified and the field is blank.'}
      </h1>
    </div>
  </div>;
