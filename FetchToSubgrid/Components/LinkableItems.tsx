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
    const records = await CrmService.getRecords(pagingFetchData, entityName);
    const entityMetadata = await CrmService.getEntityMetadata(entityName, attributesFieldNames);

    const linkEntityAttFieldNames = utilities.getLinkEntitiesNames(fetchXml ?? '');

    const items: Array<ComponentFramework.WebApi.Entity> = [];

    const linkEntityNames = Object.keys(linkEntityAttFieldNames);
    const linkEntityAttributes: any = Object.values(linkEntityAttFieldNames);

    records.entities.forEach(async (entity: any) => {
      const item: ComponentFramework.WebApi.Entity = {
        key: entity[`${entityName}id`],
        entityName,
      };

      attributesFieldNames.forEach(fieldName => {
        const type = entityMetadata.Attributes._collection[fieldName].AttributeTypeName;

        if (fieldName in entity) {
          item[fieldName] = entity[fieldName];
        }

        if (type === 'lookup' || type === 'owner') {
          item[fieldName] = <Link onClick={() => {
            CrmService.openLookupForm(entity, fieldName);
          }}>
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

      for (let i = 0; i < linkEntityNames.length; i++) {
        const linkentityMeddata: ComponentFramework.PropertyHelper.EntityMetadata =
        await CrmService.getEntityMetadata(linkEntityNames[i], linkEntityAttributes[i]);

        for (let j = 1; j < linkEntityAttributes[i].length; j++) {
          const fieldName: string = linkEntityAttributes[i][j];
          const alias: string = linkEntityAttributes[i][0];
          const type = linkentityMeddata.Attributes._collection[fieldName].AttributeTypeName;

          if (type === 'owner' || type === 'lookup') {
            item[`${alias}.${fieldName}`] = <Link onClick={() => {
              CrmService.openLinkEntityRecord(entity, `${alias}.${fieldName}`);
            }}>
              { entity[`${alias}.${fieldName}@OData.Community.Display.V1.FormattedValue`]}
            </Link>;
          }

          else {
            item[`${alias}.${fieldName}`] = entity[`${alias}.${fieldName}`];
          }
        }
      }
      items.push(item);
    });

    return items;
  },
};
