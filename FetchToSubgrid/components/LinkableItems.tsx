import { Link } from '@fluentui/react';
import * as React from 'react';
import { AttributeType } from '../utilities/enums';
import {
  openLinkEntityRecord,
  openLookupForm,
  openPrimaryEntityForm,
} from '../services/crmService';

interface ILinkableItemProps {
  item: ComponentFramework.WebApi.Entity
}

export const LinkableItem: React.FC<ILinkableItemProps> = ({ item }) => {
  if (item.isLinkEntity) {
    return (
      <Link onClick={openLinkEntityRecord.bind(null, item.entity, item.fieldName)}>
        {item.displayName}
      </Link>
    );
  }

  if (item.attributeType === AttributeType.LookUp ||
      item.attributeType === AttributeType.Owner ||
      item.attributeType === AttributeType.Customer) {
    return <Link onClick={openLookupForm.bind(null, item.entity, item.fieldName)}>
      {item.displayName}
    </Link>;
  }

  return (
    <Link onClick={openPrimaryEntityForm.bind(null, item.entity, item.entityName)}>
      {item.displayName}
    </Link>
  );
};
