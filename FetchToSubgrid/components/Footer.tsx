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
  pageSize: number,
  selectedItemsCount: number;
  totalRecordsCount: number;
  currentPage: number;
  setCurrentPage: (page: number) => void;
}

const MAX_RECORD_DISPLAY_COUNT = 5000;

export const Footer: React.FC<IFooterProps> = props => {
  const {
    pageSize,
    selectedItemsCount,
    totalRecordsCount,
    currentPage,
    setCurrentPage,
  } = props;

  const firstItemIndex = (currentPage - 1) * pageSize + 1;
  const lastItemIndex = firstItemIndex + pageSize - 1;
  const nextButtonDisabled = Math.ceil(totalRecordsCount / pageSize) <= currentPage;

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
   ${totalRecordsCount > MAX_RECORD_DISPLAY_COUNT ? '5000+' : totalRecordsCount}
   ${selectedItemsCount !== 0 ? `(${selectedItemsCount} Selected)` : ''}`;

  return (
    <div className={footerStyles.content}>
      <span>{selected}</span>
      <div className='buttons'>
        <IconButton
          styles={footerButtonStyles}
          iconProps={PreviousIcon}
          onClick={moveToFirst}
          disabled={currentPage <= 1}
        />
        <IconButton
          styles={footerButtonStyles}
          iconProps={BackIcon}
          onClick={movePrevious}
          disabled={currentPage <= 1}
        />
        <span color='black'> Page {currentPage} </span>
        <IconButton
          styles={footerButtonStyles}
          iconProps={ForwardIcon}
          onClick={moveNext}
          disabled={nextButtonDisabled}
        />
      </div>
    </div>
  );
};
