import { IButtonStyles } from '@fluentui/react/lib/components/Button/Button.types';
import { IIconProps } from '@fluentui/react/lib/components/Icon/Icon.types';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';

export const PreviousIcon: IIconProps = { iconName: 'Previous' };
export const BackIcon: IIconProps = { iconName: 'Back' };
export const ForwardIcon: IIconProps = { iconName: 'Forward' };

export const footerStyles = mergeStyleSets({
  content: {
    justifyContent: 'end !important',
    display: 'flex',
    flex: '1 1 auto',
    placeContent: 'stretch space-between',
    height: '40px',
    color: '#333',
    fontSize: '12px',
    alignItems: 'center',
  },
});

export const footerButtonStyles: Partial<IButtonStyles> = {
  root: {
    backgroundColor: 'transparent',
    cursor: 'pointer',
    height: '0px',
    color: 'green',
  },
  icon: {
    fontSize: '12px',
    backgroundColor: 'transparent',
    cursor: 'pointer',
    height: '0px',
    color: 'rgb(0 120 212)',
  },
};
