import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import {
  BackIcon,
  footerButtonStyles,
  footerStyles,
  ForwardIcon,
  PreviousIcon,
} from '../styles/footerStyles';

interface IFooterProps {
  firstItemIndex: number;
  lastItemIndex: number;
  selectedItemsCount: number;
  totalRecordsCount: number;
  currentPage: number;
  nextButtonDisable: boolean;
  movePreviousIsDisabled: boolean;
  setCurrentPage: (page: number) => void;
}

export const Footer: React.FC<IFooterProps> = props => {
  const {
    firstItemIndex,
    lastItemIndex,
    selectedItemsCount,
    totalRecordsCount,
    currentPage,
    nextButtonDisable,
    movePreviousIsDisabled,
    setCurrentPage,
  } = props;

  function moveToFirst() {
    setCurrentPage(1);
  }

  function movePrevious() {
    setCurrentPage(currentPage - 1);
  }

  function moveNext() {
    setCurrentPage(currentPage + 1);
  }

  const MAX_RECORD_DISPLAY_COUNT = 5000;

  const selected = `${firstItemIndex} - ${lastItemIndex} of 
   ${totalRecordsCount > MAX_RECORD_DISPLAY_COUNT ? '5000+' : totalRecordsCount}
   ${selectedItemsCount !== 0 ? `(${selectedItemsCount} Selected)` : ''}`;

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
