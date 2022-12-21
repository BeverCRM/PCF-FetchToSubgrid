import { Link } from '@fluentui/react';
import * as React from 'react';
import { openLinkEntityRecord, openLookupForm,
  openPrimaryEntityForm } from '../Services/CrmService';
import { AttributeType } from '../Utilities/enums';

interface IlinkableItemProps {
  item: ComponentFramework.WebApi.Entity
}

export const LinkableItem: React.FunctionComponent<IlinkableItemProps> = props => {
  const { item } = props;

  if (item.isLinkEntity) {
    return <Link onClick={() => {
      openLinkEntityRecord(item.entity, item.fieldName);
    }}>
      {item.displayName}
    </Link>;
  }

  if (item.attributeType === AttributeType.LOOKUP ||
          item.attributeType === AttributeType.OWNER) {
    return <Link onClick={() => {
      openLookupForm(item.entity, item.fieldName);
    }}>
      {item.displayName}
    </Link>;
  }

  return <Link onClick={() => {
    openPrimaryEntityForm(item.entity, item.entityName);
  }}>
    {item.displayName}
  </Link>;
};
