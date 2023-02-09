import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import {
  BackIcon,
  footerButtonStyles,
  footerStyles,
  ForwardIcon,
  PreviousIcon,
} from '../styles/footerStyles';

export interface IGridFooterProps {
  firstNumber:any;
  lastNumber:any;
  selectedItems: number;
  totalRecordsCount:number;
  currentPage: number;
  setCurrentPage: (page: number) => void;
  nextButtonDisable: boolean;
  isMovePrevious: boolean;
}

export const Footer: React.FunctionComponent<IGridFooterProps> = props => {
  const {
    firstNumber,
    lastNumber,
    selectedItems,
    totalRecordsCount,
    currentPage,
    setCurrentPage,
    nextButtonDisable,
    isMovePrevious } = props;

  function moveToFirst() {
    setCurrentPage(1);
  }

  function movePrevious() {
    setCurrentPage(currentPage - 1);
  }

  function moveNext() {
    setCurrentPage(currentPage + 1);
  }

  const selected = `${firstNumber} - ${lastNumber} of 
  ${totalRecordsCount >= 5000 ? '5000+' : totalRecordsCount}
   ${selectedItems !== 0 ? `(${selectedItems} Selected)` : ''}`;

  return <div className={footerStyles.content}>
    <span > {selected} </span>
    <div className='buttons'>
      <IconButton
        styles={footerButtonStyles}
        iconProps={PreviousIcon}
        onClick={moveToFirst}
        disabled = {isMovePrevious}
      />
      <IconButton
        styles={footerButtonStyles}
        iconProps={BackIcon}
        onClick={movePrevious}
        disabled={isMovePrevious}
      />
      <span color='black'> Page {currentPage} </span>
      <IconButton
        styles={footerButtonStyles}
        iconProps={ForwardIcon}
        onClick={moveNext}
        disabled={nextButtonDisable}
      />
    </div>
  </div>;
};
