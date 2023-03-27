import { Link } from '@fluentui/react';
import * as React from 'react';
import { Entity, IService } from '../@types/types';
import { IDataverseService } from '../services/dataverseService';
import { checkIfAttributeIsEntityReferance } from '../utilities/utils';

interface ILinkableItemProps extends IService<IDataverseService> {
  item: Entity
}

export const LinkableItem: React.FC<ILinkableItemProps> = ({
  _service: dataverseService,
  item }): JSX.Element => {
  if (item.isLinkEntity) {
    return (
      <Link onClick={() => dataverseService.openLinkEntityRecordForm(item.entity, item.fieldName)}>
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
