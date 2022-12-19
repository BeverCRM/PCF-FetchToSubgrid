import CrmService from './Services/CrmService';
// import { LinkableItems } from './Components/LinkableItems';
// import { Link } from '@fluentui/react';
// import React = require('react');

/* global HTMLCollectionOf, NodeListOf */
enum AttributeType {
  DATE_TIME = 2,
  LOOKUP = 6,
  MONEY = 8,
  OWNER = 9,
  PICKLIST = 11,
  MULTISELECT_PICKLIST = 17,
}

export default {
  addPagingToFetchXml(fetchXml: string, pageSize: number, currentPage: number) {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
    const fetch: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('fetch');

    fetch[0].setAttribute('page', `${currentPage}`);
    fetch[0].setAttribute('count', `${pageSize}`);

    return new XMLSerializer().serializeToString(xmlDoc);
  },

  getEntityName(fetchXml: string): string {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

    return xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
  },

  getLinkEntitiesNames(fetchXml: string) {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

    const linkEntities: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('link-entity');
    const linkEntityData: any = {};

    Array.prototype.slice.call(linkEntities).forEach(linkentity => {
      const entityName: string = linkentity.attributes['name'].value;
      const alias: string = linkentity.attributes['alias'].value;
      const entityAttributes: string[] = [alias];
      // eslint-disable-next-line max-len
      const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(`link-entity[name="${entityName}"] > attribute`);

      Array.prototype.slice.call(attributes).map(attr => {
        entityAttributes.push(attr.attributes.name.value);
      });

      linkEntityData[entityName] = entityAttributes;

    });

    return linkEntityData;
  },

  getAttributesFieldNames(fetchXml: string): string[] {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
    // eslint-disable-next-line max-len
    const entityName: string = xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
    const attributesFieldNames: string[] = [];

    // eslint-disable-next-line max-len
    const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(`entity[name="${entityName}"] > attribute`);

    Array.prototype.slice.call(attributes).map(attr => {
      attributesFieldNames.push(attr.attributes.name.value);
    });

    return attributesFieldNames;
  },

  async getItems(
    fetchXml: string | null,
    pageSize: number,
    currentPage: number): Promise<ComponentFramework.WebApi.Entity[]> {
    const pagingFetchData: string = fetchXml
      ? this.addPagingToFetchXml(fetchXml, pageSize, currentPage)
      : '';

    const attributesFieldNames: string[] = this.getAttributesFieldNames(pagingFetchData);
    const entityName: string = this.getEntityName(fetchXml ?? '');
    const records = await CrmService.getRecords(pagingFetchData);
    const entityMetadata = await CrmService.getEntityMetadata(entityName, attributesFieldNames);
    const linkEntityAttFieldNames = this.getLinkEntitiesNames(fetchXml ?? '');

    const items: Array<ComponentFramework.WebApi.Entity> = [];
    const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
    const linkEntityAttributes: any = Object.values(linkEntityAttFieldNames);

    const promises = linkEntityNames.map((linkEntityNames, i) => CrmService.getEntityMetadata(
      linkEntityNames, linkEntityAttributes[i]));

    const linkentityMeddataArr: ComponentFramework.PropertyHelper.EntityMetadata[] =
     await Promise.all(promises);

    records.entities.forEach((entity: any) => {
      const item: ComponentFramework.WebApi.Entity = {};

      attributesFieldNames.forEach(fieldName => {
        const attributeType = entityMetadata.Attributes._collection[fieldName].AttributeType;

        if (fieldName === entityMetadata._primaryNameAttribute) {
          item[fieldName] =
          {
            displayName: entity[fieldName],
            linkable: true,
            entity,
            fieldName,
            attributeType,
            entityName,
            isLinkEntity: false,
          };
        }

        else if (attributeType === AttributeType.MONEY ||
           attributeType === AttributeType.PICKLIST ||
           attributeType === AttributeType.DATE_TIME ||
           attributeType === AttributeType.MULTISELECT_PICKLIST) {
          item[fieldName] =
          {
            displayName: entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`],
            linkable: false,
            entity,
            fieldName,
            attributeType,
            entityName,
            isLinkEntity: false,
          };
        }

        else if (attributeType === AttributeType.LOOKUP || attributeType === AttributeType.OWNER) {
          item[fieldName] =
            {
              displayName: entity[`_${fieldName}_value@OData.Community.Display.V1.FormattedValue`],
              linkable: true,
              entity,
              fieldName,
              attributeType,
              entityName,
              isLinkEntity: false,
            };
        }

        else if (fieldName in entity) {
          item[fieldName] =
            {
              displayName: entity[fieldName],
              linkable: false,
              entity,
              fieldName,
              attributeType,
              entityName,
              isLinkEntity: false,
            };

        }
      });

      for (let i = 0; i < linkEntityNames.length; i++) {
        for (let j = 1; j < linkEntityAttributes[i].length; j++) {
          const fieldName: string = linkEntityAttributes[i][j];
          const alias: string = linkEntityAttributes[i][0];
          // eslint-disable-next-line max-len
          const attributeType = linkentityMeddataArr[i].Attributes._collection[fieldName].AttributeType;

          if (`${alias}.${fieldName}` in entity) {
            item[`${alias}.${fieldName}`] =
              {
                displayName: entity[`${alias}.${fieldName}`],
                linkable: false,
                entity,
                fieldName: `${alias}.${fieldName}`,
                attributeType,
                entityName,
                isLinkEntity: true,
              };
          }

          if (attributeType === AttributeType.LOOKUP || attributeType === AttributeType.OWNER) {
            item[`${alias}.${fieldName}`] =
              {
                displayName:
                entity[`${alias}.${fieldName}@OData.Community.Display.V1.FormattedValue`],
                linkable: true,
                entity,
                fieldName: `${alias}.${fieldName}`,
                attributeType,
                entityName,
                isLinkEntity: true,
              };
          }

          if (attributeType === AttributeType.MONEY ||
              attributeType === AttributeType.PICKLIST ||
              attributeType === AttributeType.DATE_TIME ||
              attributeType === AttributeType.MULTISELECT_PICKLIST) {

            item[`${alias}.${fieldName}`] =
            {
              displayName:
              entity[`${alias}.${fieldName}@OData.Community.Display.V1.FormattedValue`],
              linkable: false,
              entity,
              fieldName: `${alias}.${fieldName}`,
              attributeType,
              entityName,
              isLinkEntity: true,
            };
          }
        }
      }
      items.push(item);
    });

    return items;
  },
};
