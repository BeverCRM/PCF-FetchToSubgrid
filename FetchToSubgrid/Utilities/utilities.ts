/* global HTMLCollectionOf, NodeListOf */

import { getRecords, getEntityMetadata } from '../Services/CrmService';
import { AttributeType } from './enums';

export const addPagingToFetchXml =
 (fetchXml: string, pageSize: number, currentPage: number): string => {
   const parser: DOMParser = new DOMParser();
   const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
   const fetch: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('fetch');

   fetch[0].setAttribute('page', `${currentPage}`);
   fetch[0].setAttribute('count', `${pageSize}`);

   return new XMLSerializer().serializeToString(xmlDoc);
 };

export const getEntityName = (fetchXml: string): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  return xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
};

export const getLinkEntitiesNames = (fetchXml: string): Object => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const linkEntities: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('link-entity');
  const linkEntityData: { [key: string]: string[] } = {};

  Array.prototype.slice.call(linkEntities).forEach(linkentity => {
    const entityName: string = linkentity.attributes['name'].value;
    const alias: string = linkentity.attributes['alias'].value;
    const entityAttributes: string[] = [alias];

    const attributesSelector = `link-entity[name="${entityName}"] > attribute`;
    const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributesSelector);

    Array.prototype.slice.call(attributes).map(attr => {
      entityAttributes.push(attr.attributes.name.value);
    });

    linkEntityData[entityName] = entityAttributes;
  });

  return linkEntityData;
};

export const getAttributesFieldNames = (fetchXml: string): string[] => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const entityName = xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
  const attributesFieldNames: string[] = [];

  const attributeSelector = `entity[name="${entityName}"] > attribute`;
  const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributeSelector);

  Array.prototype.slice.call(attributes).map(attr => {
    attributesFieldNames.push(attr.attributes.name.value);
  });

  return attributesFieldNames;
};

const genereateItems = (
  item: ComponentFramework.WebApi.Entity, isLinkEntity: boolean,
  entityMetadata: ComponentFramework.PropertyHelper.EntityMetadata,
  attributeType: number, fieldName: string, entity: ComponentFramework.WebApi.Entity,
  entityName: string): Object => {

  let displayName: string = '';
  let linkable: boolean = false;

  if (attributeType === AttributeType.MONEY ||
      attributeType === AttributeType.PICKLIST ||
      attributeType === AttributeType.DATE_TIME ||
      attributeType === AttributeType.MULTISELECT_PICKLIST) {

    displayName = entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`];
  }

  else if (isLinkEntity) {
    if (attributeType === AttributeType.LOOKUP || attributeType === AttributeType.OWNER) {
      displayName = entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`];
      linkable = true;
    }

    else if (fieldName in entity) {
      displayName = entity[fieldName];
    }
  }

  else if (fieldName === entityMetadata._primaryNameAttribute) {
    displayName = entity[fieldName];
    linkable = true;
  }

  else if (attributeType === AttributeType.LOOKUP || attributeType === AttributeType.OWNER) {
    displayName = entity[`_${fieldName}_value@OData.Community.Display.V1.FormattedValue`];
    linkable = true;
  }

  else if (fieldName in entity) {
    displayName = entity[fieldName];
  }

  return item[fieldName] = {
    displayName,
    linkable,
    entity,
    fieldName,
    attributeType,
    entityName,
    isLinkEntity,
  };
};

export const getItems = async (
  fetchXml: string | null,
  pageSize: number,
  currentPage: number): Promise<ComponentFramework.WebApi.Entity[]> => {
  const pagingFetchData: string = fetchXml
    ? addPagingToFetchXml(fetchXml, pageSize, currentPage)
    : '';

  const attributesFieldNames: string[] = getAttributesFieldNames(pagingFetchData);
  const entityName: string = getEntityName(fetchXml ?? '');
  const records = await getRecords(pagingFetchData);
  const entityMetadata = await getEntityMetadata(entityName, attributesFieldNames);

  const linkEntityAttFieldNames = getLinkEntitiesNames(fetchXml ?? '');
  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes: any = Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => getEntityMetadata(
    linkEntityNames, linkEntityAttributes[i]));

  const linkentityMetadata: ComponentFramework.PropertyHelper.EntityMetadata[] =
     await Promise.all(promises);

  const items: ComponentFramework.WebApi.Entity[] = [];

  records.entities.forEach(entity => {
    const item: ComponentFramework.WebApi.Entity = {
      id: entity[`${entityName}id`],
    };

    attributesFieldNames.forEach(fieldName => {
      const attributeType = entityMetadata.Attributes._collection[fieldName].AttributeType;

      genereateItems(
        item,
        false,
        entityMetadata,
        attributeType,
        fieldName,
        entity,
        entityName);
    });

    linkEntityAttributes.forEach(([alias, ...fields]: [string, string], i: number) => {
      fields.forEach(fieldName => {
        const attributeType = linkentityMetadata[i].Attributes._collection[fieldName].AttributeType;

        genereateItems(item, true, linkentityMetadata,
          attributeType, `${alias}.${fieldName}`, entity, entityName);
      });
    });
    items.push(item);
  });

  return items;
};
