import { CommandBarButton, IIconProps, IStackStyles } from '@fluentui/react';
import * as React from 'react';
import { openRecord, openRecordDeleteDialog } from '../services/crmService';
import { ContainerButtonStyles } from '../styles/comandBarStyles';

export interface ICommandBarProps {
  className: string;
  entityName: string;
  selectedRecordIds: string[];
  displayName: string;
  setDialogAccepted: React.Dispatch<React.SetStateAction<boolean>>;
  newButtonVisibility: boolean;
  deleteButtonVisibility: boolean;
}

export const stackStyles: Partial<IStackStyles> = { root: { height: 44, marginLeft: 100 } };
const deleteIcon: IIconProps = { iconName: 'Delete' };
const addIcon: IIconProps = { iconName: 'Add' };

export const CommandBar = ({
  className,
  entityName,
  selectedRecordIds,
  displayName,
  setDialogAccepted,
  newButtonVisibility,
  deleteButtonVisibility } : ICommandBarProps) =>
  <div className='containerButtons'>
    <CommandBarButton
      maxLength={1}
      disabled={!newButtonVisibility}
      iconProps={addIcon}
      styles={ContainerButtonStyles}
      text={`New ${displayName}`}
      onClick={() => openRecord(entityName, '')}
    />

    <CommandBarButton
      className={className}
      disabled={!deleteButtonVisibility}
      iconProps={deleteIcon}
      styles={ContainerButtonStyles}
      text="Delete"
      onClick={() => openRecordDeleteDialog(selectedRecordIds, entityName, setDialogAccepted) }
    />
  </div>;
