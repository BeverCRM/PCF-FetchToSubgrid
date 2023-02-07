import { CommandBarButton, IIconProps, IStackStyles } from '@fluentui/react';
import React = require('react');
import { openRecord, openRecordDeleteDialog } from '../services/crmService';
import { ContainerButtonStyles } from '../styles/comandBarStyles';

export interface ICommandBarProps {
  className: string;
  entityName: string;
  selectedRecordIds: string[];
  displayName: string;
  setIsUsedButton: any;
  newButtonVisibility: string | null;
  deleteButtonVisibility: string | null;
}

export const stackStyles: Partial<IStackStyles> = { root: { height: 44, marginLeft: 100 } };
const deleteIcon: IIconProps = { iconName: 'Delete' };
const addIcon: IIconProps = { iconName: 'Add' };

export const CommandBar = ({
  className,
  entityName,
  selectedRecordIds,
  displayName,
  setIsUsedButton,
  newButtonVisibility,
  deleteButtonVisibility } : ICommandBarProps) =>
  <div className='containerButtons'>
    <CommandBarButton
      maxLength={1}
      disabled={newButtonVisibility !== 'true'}
      iconProps={addIcon}
      styles={ContainerButtonStyles}
      text={`New ${displayName}`}
      onClick={() => { openRecord(entityName, ''); }}
    />

    <CommandBarButton
      className={className}
      disabled={deleteButtonVisibility !== 'true'}
      iconProps={deleteIcon}
      styles={ContainerButtonStyles}
      text="Delete"
      onClick={() => {
        openRecordDeleteDialog(
          selectedRecordIds,
          entityName,
          setIsUsedButton);
      }}
    />
  </div>;
