import { IButtonStyles } from '@fluentui/react/lib/components/Button/Button.types';
import { IIconProps } from '@fluentui/react/lib/components/Icon/Icon.types';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';

export const deleteIcon: IIconProps = { iconName: 'Delete' };
export const downloadIcon: IIconProps = { iconName: 'Download' };
export const refreshIcon: IIconProps = { iconName: 'Refresh' };
export const addIcon: IIconProps = { iconName: 'Add' };

export const dataSetStyles = mergeStyleSets({
  content: {
    width: '100%',
    display: 'inline-block',
    position: 'relative',
    height: '40px',
  },
  buttons: {
    height: '44px',
    paddingRight: '20px',
  },
  detailsList: {
    paddingTop: '0px',
  },
  commandBarButton: {
    root: {
      color: 'black',
    },
    icon: {
      color: 'black',
    },
  },
});

export const ContainerButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: 'rgb(51, 51, 51)',
    backgroundColor: 'white',
    height: 43,
  },
  icon: {
    color: 'rgb(51, 51, 51);',
  },
};

export const CommandBarButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: 'black',
    backgroundColor: 'white',
  },
  icon: {
    color: 'black',
  },
};
