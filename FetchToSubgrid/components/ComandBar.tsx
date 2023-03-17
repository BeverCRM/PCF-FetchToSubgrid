import { CommandBarButton, IIconProps } from '@fluentui/react';
import * as React from 'react';
import { ContainerButtonStyles } from '../styles/comandBarStyles';
import { ICommandBarProps } from '../@types/types';

const deleteIcon: IIconProps = { iconName: 'Delete' };
const addIcon: IIconProps = { iconName: 'Add' };

export const CommandBar = ({
  _service: dataverseService,
  isButtonActive,
  entityName,
  selectedRecordIds,
  displayName,
  newButtonVisibility,
  deleteButtonVisibility,
  setDialogAccepted }: ICommandBarProps) =>
  <div className='containerButtons'>
    {newButtonVisibility &&
        <CommandBarButton
          styles={ContainerButtonStyles}
          maxLength={1}
          iconProps={addIcon}
          text={`New ${displayName}`}
          onClick={() => dataverseService.openRecord(entityName, '')}
        />
    }
    {deleteButtonVisibility && isButtonActive &&
        <CommandBarButton
          styles={ContainerButtonStyles}
          iconProps={deleteIcon}
          text="Delete"
          onClick={() => dataverseService.openRecordDeleteDialog(
            selectedRecordIds, entityName, setDialogAccepted)}
        />
    }
  </div>;
