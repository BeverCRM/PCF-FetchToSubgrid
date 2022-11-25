import { Link } from '@fluentui/react';
import * as React from 'react';
import CrmService from '../Services/CrmService';
import utilities from '../utilities';

export default {

  async getLinkableItems(fetchXml: string | null): Promise<ComponentFramework.WebApi.Entity[]> {
    const attributesFieldNames: string[] = utilities.getAttributesFieldNames(fetchXml ?? '');
    const entityName: string = utilities.getEntityName(fetchXml ?? '');

    const record = await CrmService.getRecord(fetchXml, entityName);
    const entityMetadata = await CrmService.getEntityMetadata(entityName, attributesFieldNames);

    const items: Array<ComponentFramework.WebApi.Entity> = [];

    record.entities.forEach(entity => {
      const item: ComponentFramework.WebApi.Entity = {};

      attributesFieldNames.forEach(fieldName => {
        if (fieldName in entity) {
          item[fieldName] = entity[fieldName];
        }

        if (`_${fieldName}_value` in entity) {
          item[fieldName] = <Link onClick={() => CrmService.openLookupForm(entity, fieldName) }>
            {entity[`_${fieldName}_value@OData.Community.Display.V1.FormattedValue`]}
          </Link>;
        }

        if (fieldName === entityMetadata._primaryNameAttribute) {
          item[fieldName] = <Link onClick={() => {
            CrmService.openPrimaryEntityForm(entity, entityName);

          }}>
            {entity[fieldName]}
          </Link>;
        }
      });

      items.push(item);
    });

    return items;
  },
};
