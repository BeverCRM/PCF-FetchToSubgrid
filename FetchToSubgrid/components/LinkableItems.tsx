import { Link } from '@fluentui/react';
import * as React from 'react';
import { Entity, IService } from '../@types/types';
import { IDataverseService } from '../services/dataverseService';
import { checkIfAttributeIsEntityReferance } from '../utilities/utils';

interface ILinkableItemProps extends IService<IDataverseService> {
  item: Entity
}

export const LinkableItem: React.FC<ILinkableItemProps> = ({
  _service: dataverseService, item }): JSX.Element => {
  const onLinkClick = () => {
    if (item.isLinkEntity) {
      dataverseService.openLinkEntityRecordForm(item.entity, item.fieldName);
    }
    else if (checkIfAttributeIsEntityReferance(item.AttributeType)) {
      dataverseService.openLookupForm(item.entity, item.fieldName);
    }
    else {
      dataverseService.openPrimaryEntityForm(item.entity, item.entityName);
    }
  };

  return (
    <Link onClick={onLinkClick}>
      {item.displayName}
    </Link>
  );
};
