import * as React from 'react';
import { IFooterProps } from '../@types/types';
import { IconButton } from '@fluentui/react/lib/Button';
import {
  BackIcon,
  footerButtonStyles,
  footerStyles,
  ForwardIcon,
  PreviousIcon,
} from '../styles/footerStyles';

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

  const selected = `${firstItemIndex} - ${lastItemIndex} of 
   ${totalRecordsCount >= 5000 ? '5000+' : totalRecordsCount}
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
