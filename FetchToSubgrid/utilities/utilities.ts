/* global HTMLCollectionOf, NodeListOf */

import { getRecords, getEntityMetadata } from '../services/crmService';
import { AttributeType } from './enums';
import { Dictionary, Entity, EntityMetadata, EntityMetadataDictionary, IItemProps } from './types';

export const addPagingToFetchXml = (
  fetchXml: string,
  pageSize: number,
  currentPage: number): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];
  const top: string | null = fetch.getAttribute('top');

  if (Number(top)) {
    fetch.removeAttribute('top');
    fetch.setAttribute('page', `${currentPage}`);
    fetch.setAttribute('count', `${top}`);
  }
  else {
    fetch.setAttribute('page', `${currentPage}`);
    fetch.setAttribute('count', `${pageSize}`);
  }

  return new XMLSerializer().serializeToString(xmlDoc);
};

export const getEntityName = (fetchXml: string): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  return xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
};

export const getLinkEntitiesNames = (fetchXml: string): Dictionary => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const linkEntities: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('link-entity');
  const linkEntityData: Dictionary = {};

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

export const getAliasNames = (fetchXml: string): string[] => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const aggregateAttrNames: string[] = [];

  const entityName = xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
  const attributeSelector = `entity[name="${entityName}"] > attribute`;
  const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributeSelector);

  Array.prototype.slice.call(attributes).map(attr => {
    aggregateAttrNames.push(attr.attributes.alias.value);
  });

  return aggregateAttrNames;
};

export const isAggregate = (fetchXml: string): boolean => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const aggregate = xmlDoc.getElementsByTagName('fetch')?.[0]?.getAttribute('aggregate') ?? '';
  if (aggregate === 'true') return true;

  return false;
};

const genereateItems = (props: IItemProps): Entity => {
  const {
    item,
    isLinkEntity,
    entityMetadata,
    attributeType,
    fieldName,
    entity,
    fetchXml,
    index,
  } = props;

  let displayName = '';
  let linkable = false;

  const hasAggregate: boolean = isAggregate(fetchXml ?? '');
  const entityName: string = getEntityName(fetchXml ?? '');

  if (hasAggregate) {
    const aggregateAttrNames = getAliasNames(fetchXml ?? '');

    return item[aggregateAttrNames[index]] = {
      displayName: entity[aggregateAttrNames[index]],
      linkable: false,
      entity,
      fieldName: aggregateAttrNames[index],
      attributeType,
      entityName,
      isLinkEntity,
      aggregate: true,
    };
  }

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

export const getCountInFetchXml = (fetchXml: string | null): number => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');
  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

  const count: string | null = fetch.getAttribute('count');
  const top: string | null = fetch.getAttribute('top');

  return Number(count) || Number(top);
};

export const getItems = async (
  fetchXml: string | null,
  pageSize: number,
  currentPage: number): Promise<Entity[]> => {

  const countOfRecordsInFetch = getCountInFetchXml(fetchXml);

  const pagingFetchData: string = addPagingToFetchXml(
    fetchXml ?? '',
    countOfRecordsInFetch || pageSize,
    currentPage);

  let attributesFieldNames: string[] = getAttributesFieldNames(pagingFetchData);
  const entityName: string = getEntityName(fetchXml ?? '');
  const records: ComponentFramework.WebApi.RetrieveMultipleResponse = await getRecords(
    pagingFetchData);

  let isAllAttribute = false;

  if (attributesFieldNames.length === 0 && entityName) {
    attributesFieldNames = Object.keys(records.entities[0]);
    isAllAttribute = true;
  }

  const entityMetadata: EntityMetadata = await getEntityMetadata(entityName, attributesFieldNames);
  const attributesCollection: EntityMetadataDictionary = entityMetadata.Attributes._collection;

  if (isAllAttribute) {
    attributesFieldNames = Object.keys(attributesCollection);
  }

  const linkEntityAttFieldNames: Dictionary = getLinkEntitiesNames(fetchXml ?? '');
  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes: any = Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => getEntityMetadata(
    linkEntityNames, linkEntityAttributes[i]));

  const linkentityMetadata: EntityMetadata[] = await Promise.all(promises);

  const items: Entity[] = [];

  records.entities.forEach(entity => {
    const item: Entity = {
      id: entity[`${entityName}id`],
    };

    attributesFieldNames.forEach((fieldName, index) => {
      const attributeType: number = entityMetadata.Attributes._collection[fieldName].AttributeType;

      const attributes = {
        item,
        isLinkEntity: false,
        entityMetadata,
        attributeType,
        fieldName,
        entity,
        fetchXml,
        index,
      };

      genereateItems(attributes);
    });

    linkEntityAttributes.forEach(([alias, ...fields]: [string, string], index: number) => {
      fields.forEach(fieldName => {
        const attributeType: number =
         linkentityMetadata[index].Attributes._collection[fieldName].AttributeType;

        const attributes = {
          item,
          isLinkEntity: true,
          entityMetadata: linkentityMetadata[index],
          attributeType,
          fieldName: `${alias}.${fieldName}`,
          entity,
          fetchXml,
          index,
        };
        genereateItems(attributes);
      });
    });
    items.push(item);
  });

  return items;
};
