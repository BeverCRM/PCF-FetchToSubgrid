import { CommandBarButton, IIconProps } from '@fluentui/react';
import * as React from 'react';
import { IService } from '../@types/types';
import { IDataverseService } from '../services/dataverseService';
import { ContainerButtonStyles } from '../styles/comandBarStyles';

const deleteIcon: IIconProps = { iconName: 'Delete' };
const addIcon: IIconProps = { iconName: 'Add' };

interface ICommandBarProps extends IService<IDataverseService> {
  isButtonActive: boolean;
  entityName: string;
  selectedRecordIds: string[];
  newButtonVisibility: boolean;
  deleteButtonVisibility: boolean | string;
  setDialogAccepted: React.Dispatch<React.SetStateAction<boolean>>;
}

export const CommandBar = ({
  _service: dataverseService,
  isButtonActive,
  entityName,
  selectedRecordIds,
  newButtonVisibility,
  deleteButtonVisibility,
  setDialogAccepted }: ICommandBarProps) => {
  const displayName = React.useRef('');

  React.useEffect(() => {
    const fetchDisplayName = async () => {
      displayName.current = await dataverseService.getEntityDisplayName(entityName);
    };
    fetchDisplayName();
  }, [entityName]);

  const handleNewButtonClick = async () => {
    dataverseService.openNewRecord(entityName);
  };

  const handleDeleteButtonClick = async () => {
    const deleteDialogStatus = await dataverseService.openRecordDeleteDialog(entityName);

    if (deleteDialogStatus.confirmed) {
      setDialogAccepted(true);
      await dataverseService.deleteSelectedRecords(selectedRecordIds, entityName);
      setDialogAccepted(false);
    }
  };

  return (
    <div className='containerButtons'>
      {newButtonVisibility &&
        <CommandBarButton
          styles={ContainerButtonStyles}
          maxLength={1}
          iconProps={addIcon}
          text={`New ${displayName.current}`}
          onClick={handleNewButtonClick}
        />
      }
      {deleteButtonVisibility && isButtonActive &&
        <CommandBarButton
          styles={ContainerButtonStyles}
          iconProps={deleteIcon}
          text="Delete"
          onClick={handleDeleteButtonClick}
        />
      }
    </div>);
};
