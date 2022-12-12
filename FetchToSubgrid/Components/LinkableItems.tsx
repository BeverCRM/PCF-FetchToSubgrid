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
        const attributeType = entityMetadata.Attributes._collection[fieldName].AttributeType;

        if (fieldName in entity) {
          item[fieldName] = entity[fieldName];
        }
        // money-8, picklist-11, datatime-2 multiselectpicklist-17
        if (attributeType === 8 || attributeType === 11 ||
          attributeType === 2 || attributeType === 17) {

          item[fieldName] =
            entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`];
        }
        // lookup-6 owner-9
        if (attributeType === 6 || attributeType === 9) {
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
          const attributeType = linkentityMeddata.Attributes._collection[fieldName].AttributeType;

          if (`${alias}.${fieldName}` in entity) {
            item[`${alias}.${fieldName}`] = entity[`${alias}.${fieldName}`];
          }
          // lookup-6 owner-9
          if (attributeType === 6 || attributeType === 9) {
            item[`${alias}.${fieldName}`] = <Link onClick={() => {
              CrmService.openLinkEntityRecord(entity, `${alias}.${fieldName}`);
            }}>
              { entity[`${alias}.${fieldName}@OData.Community.Display.V1.FormattedValue`]}
            </Link>;
          }
          // money-8, picklist-11, datatime-2 multiselectpicklist-17
          if (attributeType === 8 || attributeType === 11 ||
          attributeType === 2 || attributeType === 17) {

            item[`${alias}.${fieldName}`] =
             entity[`${alias}.${fieldName}@OData.Community.Display.V1.FormattedValue`];
          }
        }
      }
      items.push(item);
    });

    return items;
  },
};
