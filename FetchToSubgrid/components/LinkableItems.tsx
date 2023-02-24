import { Link } from '@fluentui/react';
import * as React from 'react';
import { ILinkableItemProps } from '../utilities/types';
import { checkIfAttributeIsEntityReferance } from '../utilities/utils';

export const LinkableItem: React.FC<ILinkableItemProps> = ({
  _service: dataverseService,
  item }): JSX.Element => {
  if (item.isLinkEntity) {
    return (
      <Link onClick={() => dataverseService.openLinkEntityRecord(item.entity, item.fieldName)}>
        {item.displayName}
      </Link>
    );
  }

  if (checkIfAttributeIsEntityReferance(item.AttributeType)) {
    return (
      <Link onClick={() => dataverseService.openLookupForm(item.entity, item.fieldName)}>
        {item.displayName}
      </Link>
    );
  }

  return (
    <Link onClick={() => dataverseService.openPrimaryEntityForm(item.entity, item.entityName)}>
      {item.displayName}
    </Link>
  );
};
