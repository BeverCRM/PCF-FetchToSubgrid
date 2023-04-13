import * as React from 'react';
import { IDataverseService } from '../services/dataverseService';

interface IErrorMessageProps {
  error: Error;
  dataverseService: IDataverseService;
}

export const ErrorMessage: React.FC<IErrorMessageProps> = ({ error, dataverseService }) =>
  <div className='errorMessage'>
    <h1
      className='errorMessageText'
      onClick={() => dataverseService.openErrorDialog(error)}
    >
      An error has occurred!
    </h1>
  </div>;
