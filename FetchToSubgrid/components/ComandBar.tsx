import { CommandBarButton, IIconProps, IStackStyles } from '@fluentui/react';
import React = require('react');
import { openRecord, openRecordDeleteDialog } from '../services/crmService';
import { ContainerButtonStyles } from '../styles/comandBarStyles';

export interface ICommandBarProps {
  entityName: string;
  selectedRecordIds: string[];
  displayName: string;
  setIsUsedButton: any;
  userParameters: any;
}

export const stackStyles: Partial<IStackStyles> = { root: { height: 44, marginLeft: 100 } };
const deleteIcon: IIconProps = { iconName: 'Delete' };
const addIcon: IIconProps = { iconName: 'Add' };

export const CommandBar = ({
  entityName,
  selectedRecordIds,
  displayName,
  setIsUsedButton,
  userParameters } : ICommandBarProps) =>
  <div className='containerButtons'>
    <CommandBarButton
      maxLength={1}
      disabled = { !(userParameters.NewVIsiblitiy === 'true') }
      iconProps={addIcon}
      styles={ContainerButtonStyles}
      text={`New ${displayName}`}
      onClick={() => { openRecord(entityName, ''); }}
    />

    <CommandBarButton
      disabled = { !(userParameters.DeleteVisiblity === 'true') }
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
  </div>
  ;
