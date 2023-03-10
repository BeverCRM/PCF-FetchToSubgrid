/* global HTMLCollectionOf, NodeListOf */
import { IColumn } from '@fluentui/react';
import { Dictionary, EntityAttribute } from './types';

export const changeAliasNames = (fetchXml: string) => {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(fetchXml, 'text/xml');

  const attributeElements = xmlDoc.querySelectorAll('attribute[alias]');

  attributeElements.forEach((attributeElement, index) => {
    const newAliasValue = `alias${index}`;
    attributeElement.setAttribute('alias', newAliasValue);
  });

  return xmlDoc.documentElement.outerHTML;
};

export const addPagingToFetchXml =
 (fetchXml: string, pageSize: number, currentPage :number, recordsCount: number): string => {
   const parser: DOMParser = new DOMParser();
   const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

   const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];
   fetch?.removeAttribute('count');
   fetch?.removeAttribute('page');
   fetch?.removeAttribute('top');

   let recordsPerPage = pageSize;

   if (currentPage * recordsPerPage > recordsCount) {
     recordsPerPage = recordsCount % pageSize;
   }

   fetch.setAttribute('page', `${currentPage}`);
   fetch.setAttribute('count', `${recordsPerPage}`);

   const newFetchChangedAliases = changeAliasNames(new XMLSerializer().serializeToString(xmlDoc));

   return newFetchChangedAliases;
 };

export const getEntityNameFromFetchXml = (fetchXml: string): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  return xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
};

export const getOrderInFetch = (fetchXml: string) => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: XMLDocument = parser.parseFromString(fetchXml, 'text/xml');
  const entityOrder: any = xmlDoc.querySelector('entity > order');
  const linkEntityOrder: any = xmlDoc.querySelectorAll('link-entity > order');

  if (entityOrder && linkEntityOrder.length > 0 || linkEntityOrder.length > 1) return null;

  const order = entityOrder || linkEntityOrder[0];
  if (!order) return null;

  const isLinkEntity = linkEntityOrder[0] !== null;

  const descending: boolean = order.attributes.descending?.value === 'true';
  const attribute: string = order.attributes.attribute?.value;

  return {
    [attribute]: descending,
    isLinkEntity,
  };
};

export const addOrderToFetch = (
  fetchXml: string,
  fieldName: string,
  column?: IColumn,
): string => {

  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const entity: Element = xmlDoc.getElementsByTagName('entity')[0];
  const linkEntity: Element = xmlDoc.getElementsByTagName('link-entity')[0];

  const entityOrder: Element = entity?.getElementsByTagName('order')[0];
  const linkOrder: Element = linkEntity?.getElementsByTagName('order')[0];

  if (entityOrder) {
    const parent: Element = linkOrder ? linkEntity : entity;
    parent.removeChild(linkOrder || entityOrder);
  }

  if (column?.className === 'linkEntity') {
    const newOrder: HTMLElement = xmlDoc.createElement('order');
    newOrder.setAttribute('attribute', `${fieldName}`);
    newOrder.setAttribute('descending', `${!column.isSortedDescending}`);
    linkEntity.appendChild(newOrder);
  }
  else {
    const newOrder: HTMLElement = xmlDoc.createElement('order');
    newOrder.setAttribute('attribute', `${fieldName}`);
    newOrder.setAttribute('descending', `${!column?.isSortedDescending}`);
    entity.appendChild(newOrder);
  }

  return new XMLSerializer().serializeToString(xmlDoc);
};

export const getLinkEntitiesNamesFromFetchXml = (
  fetchXml: string): Dictionary<EntityAttribute[]> => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const linkEntities: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('link-entity');
  const linkEntityData: Dictionary<EntityAttribute[]> = {};

  Array.prototype.slice.call(linkEntities).forEach(linkentity => {
    const entityName: string = linkentity.attributes['name'].value;
    const entityAttributes: EntityAttribute[] = [];

    const attributesSelector = `link-entity[name="${entityName}"] > attribute`;
    const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributesSelector);
    const linkEntityAlias: string | undefined = linkentity.attributes['alias']?.value;

    Array.prototype.slice.call(attributes).map(attr => {
      const attributeAlias: string = attr.attributes['alias']?.value ?? '';

      entityAttributes.push({
        linkEntityAlias,
        name: attr.attributes.name.value,
        attributeAlias,
      });
    });

    linkEntityData[entityName] = entityAttributes;
  });

  return linkEntityData;
};

export const getAttributesFieldNamesFromFetchXml = (fetchXml: string): string[] => {
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

export const getEntityAggregateAliasNames = (fetchXml: string): string[] => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const aggregateAttrNames: string[] = [];

  const entityName = xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
  const attributeSelector = `entity[name="${entityName}"] > attribute`;
  const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributeSelector);

  Array.prototype.slice.call(attributes).map(attr => {
    aggregateAttrNames.push(attr.attributes.alias?.value);
  });

  return aggregateAttrNames;
};

export const getLinkEntityAggregateAliasNames = (fetchXml: string, i: number): string[] => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const aggregateAttrNames: string[] = [];

  const entityName = xmlDoc.getElementsByTagName('link-entity')?.[i]?.getAttribute('name') ?? '';
  const attributeSelector = `link-entity[name="${entityName}"] > attribute`;
  const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributeSelector);

  Array.prototype.slice.call(attributes).map(attr => {
    aggregateAttrNames.push(attr.attributes.alias?.value);
  });

  return aggregateAttrNames;
};

export const isAggregate = (fetchXml: string | null): boolean => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const aggregate = xmlDoc.getElementsByTagName('fetch')?.[0]?.getAttribute('aggregate') ?? '';
  if (aggregate === 'true') return true;

  return false;
};

export const getTopInFetchXml = (fetchXml: string | null): number => {
  if (!fetchXml) return 0;
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');
  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

  const top: string | null = fetch?.getAttribute('top');

  return Number(top) || 0;
};

export const getFetchXmlParserError = (fetchXml: string | null): string | null => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const parserError: Element | null = xmlDoc.querySelector('parsererror');
  if (!parserError) return null;

  const errorMessage: string | null = parserError.querySelector('div')?.innerText ?? null;
  return errorMessage;
};
