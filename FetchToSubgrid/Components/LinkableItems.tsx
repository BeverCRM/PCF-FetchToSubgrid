import { Link } from '@fluentui/react';
import * as React from 'react';
import CrmService from '../Services/CrmService';

enum AttributeType {
  LOOKUP = 6,
  OWNER = 9,
}

export default {
  makeItemsLinkable(item: any) {
    if (item.isLinkEntity) {
      return <Link onClick={() => {
        CrmService.openLinkEntityRecord(item.entity, item.fieldName);
      }}>
        {item.displayName}
      </Link>;
    }

    if (item.attributeType === AttributeType.LOOKUP ||
          item.attributeType === AttributeType.OWNER) {
      return <Link onClick={() => {
        CrmService.openLookupForm(item.entity, item.fieldName);
      }}>
        {item.displayName}
      </Link>;
    }

    return <Link onClick={() => {
      CrmService.openPrimaryEntityForm(item.entity, item.entityName);
    }}>
      {item.displayName}
    </Link>;
  },
};
