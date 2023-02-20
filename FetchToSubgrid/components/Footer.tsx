import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import {
  BackIcon,
  footerButtonStyles,
  footerStyles,
  ForwardIcon,
  PreviousIcon,
} from '../styles/footerStyles';

export interface IFooterProps {
  firstItemIndex: number;
  lastItemIndex: number;
  selectedItems: number;
  totalRecordsCount:number;
  currentPage: number;
  setCurrentPage: (page: number) => void;
  nextButtonDisable: boolean;
  movePreviousIsDisabled: boolean;
}

export const Footer: React.FC<IFooterProps> = props => {
  const {
    firstItemIndex,
    lastItemIndex,
    selectedItems,
    totalRecordsCount,
    currentPage,
    setCurrentPage,
    nextButtonDisable,
    movePreviousIsDisabled } = props;

  function moveToFirst() {
    setCurrentPage(1);
  }

  function movePrevious() {
    setCurrentPage(currentPage - 1);
  }

  function moveNext() {
    setCurrentPage(currentPage + 1);
  }

  const selected = `${firstItemIndex} - ${lastItemIndex} of 
   ${totalRecordsCount >= 5000 ? '5000+' : totalRecordsCount}
   ${selectedItems !== 0 ? `(${selectedItems} Selected)` : ''}`;

  return (
    <div className={footerStyles.content}>
      <span > {selected} </span>
      <div className='buttons'>
        <IconButton
          styles={footerButtonStyles}
          iconProps={PreviousIcon}
          onClick={moveToFirst}
          disabled = {movePreviousIsDisabled}
        />
        <IconButton
          styles={footerButtonStyles}
          iconProps={BackIcon}
          onClick={movePrevious}
          disabled={movePreviousIsDisabled}
        />
        <span color='black'> Page {currentPage} </span>
        <IconButton
          styles={footerButtonStyles}
          iconProps={ForwardIcon}
          onClick={moveNext}
          disabled={nextButtonDisable}
        />
      </div>
    </div>
  );
};
