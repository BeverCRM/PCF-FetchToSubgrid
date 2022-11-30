import { Link } from '@fluentui/react';
import * as React from 'react';
import CrmService from '../Services/CrmService';
import utilities from '../utilities';

export default {

  async getLinkableItems(fetchXml: string | null,
    pageSize: number,
    currentPage: number,
  ): Promise<ComponentFramework.WebApi.Entity[]> {

    const pagingFetchData: string = fetchXml
      ? utilities.addPagingToFetchXml(fetchXml, pageSize, currentPage) : '';

    const attributesFieldNames: string[] = utilities.getAttributesFieldNames(pagingFetchData);
    const entityName: string = utilities.getEntityName(fetchXml ?? '');

    const record = await CrmService.getRecord(pagingFetchData, entityName);
    const entityMetadata = await CrmService.getEntityMetadata(entityName, attributesFieldNames);

    const items: Array<ComponentFramework.WebApi.Entity> = [];

    record.entities.forEach(entity => {
      const item: ComponentFramework.WebApi.Entity = {
        key: entity[`${entityName}id`],
        entityName,
      };

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
