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
  currentPage: number;
  setCurrentPage: (page: number) => void;
  nextButtonDisable: boolean;
  isMovePrevious: boolean;
}

export const Footer: React.FunctionComponent<IGridFooterProps> = props => {
  const { currentPage, setCurrentPage, nextButtonDisable, isMovePrevious } = props;

  function moveToFirst() {
    setCurrentPage(1);
  }

  function movePrevious() {
    setCurrentPage(currentPage - 1);
  }

  function moveNext() {
    setCurrentPage(currentPage + 1);
  }

  return <div className={footerStyles.content}>
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
