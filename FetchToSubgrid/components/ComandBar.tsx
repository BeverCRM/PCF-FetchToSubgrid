import { CommandBarButton, IIconProps, IStackStyles } from '@fluentui/react';
import * as React from 'react';
import { ContainerButtonStyles } from '../styles/comandBarStyles';
import { ICommandBarProps } from '../utilities/types';

export const stackStyles: Partial<IStackStyles> = { root: { height: 44, marginLeft: 100 } };
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
  setDialogAccepted }: ICommandBarProps) => {

  const isButtonVisible = deleteButtonVisibility && isButtonActive;

  return (
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

      {isButtonVisible &&
        <CommandBarButton
          styles={ContainerButtonStyles}
          iconProps={deleteIcon}
          text="Delete"
          onClick={() => dataverseService.openRecordDeleteDialog(
            selectedRecordIds, entityName, setDialogAccepted)}
        />
      }
    </div>
  );
};
