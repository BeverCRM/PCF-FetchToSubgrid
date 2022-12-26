import { Link } from '@fluentui/react';
import * as React from 'react';
import { AttributeType } from '../Utilities/enums';
import {
  openLinkEntityRecord,
  openLookupForm,
  openPrimaryEntityForm,
} from '../Services/CrmService';

interface IlinkableItemProps {
  item: ComponentFramework.WebApi.Entity
}

export const LinkableItem: React.FunctionComponent<IlinkableItemProps> = props => {
  const { item } = props;

  if (item.isLinkEntity) {
    return <Link onClick={openLinkEntityRecord.bind(null, item.entity, item.fieldName)}>
      {item.displayName}
    </Link>;
  }

  if (item.attributeType === AttributeType.LOOKUP || item.attributeType === AttributeType.OWNER) {
    return <Link onClick={openLookupForm.bind(null, item.entity, item.fieldName)}>
      {item.displayName}
    </Link>;
  }

  return <Link onClick={openPrimaryEntityForm.bind(null, item.entity, item.entityName)}>
    {item.displayName}
  </Link>;
};
