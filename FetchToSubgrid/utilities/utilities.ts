/* global HTMLCollectionOf, NodeListOf */
import * as React from 'react';
import { IColumn } from '@fluentui/react';
import { AttributeType } from './enums';
import {
  getCurrentPageRecords,
  getEntityMetadata,
  getWholeNumberFieldName,
  getTimeZoneDefinitions,
  getRecordsCount } from '../services/crmService';
import {
  Entity,
  EntityMetadata,
  Dictionary,
  RetriveRecords,
  IItemProps,
  EntityAttribute } from './types';

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

   return new XMLSerializer().serializeToString(xmlDoc);
 };

export const getEntityName = (fetchXml: string): string => {
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

  const descending: string = order.attributes.descending.value;
  const attribute: string = order.attributes.attribute.value;

  return {
    [descending]: attribute,
    isLinkEntity,
  };
};

export const addOrderToFetch = (fetchXml: string,
  fieldName: string,
  dialogEvent: any,
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
    newOrder.setAttribute('descending', `${dialogEvent.key === 'zToA'}`);
    linkEntity.appendChild(newOrder);
  }
  else {
    const newOrder: HTMLElement = xmlDoc.createElement('order');
    newOrder.setAttribute('attribute', `${fieldName}`);
    newOrder.setAttribute('descending', `${dialogEvent.key === 'zToA'}`);
    entity.appendChild(newOrder);
  }

  return new XMLSerializer().serializeToString(xmlDoc);
};

export const getLinkEntitiesNames = (fetchXml: string): Dictionary<EntityAttribute[]> => {
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

const getEntityData = (props: IItemProps) => {
  const {
    timeZoneDefinitions,
    isLinkEntity,
    entityMetadata,
    attributeType,
    fieldName,
    entity,
  } = props;

  const needToGetFormattedValue1 = () =>
    attributeType === AttributeType.Money ||
    attributeType === AttributeType.PickList ||
    attributeType === AttributeType.DateTime ||
    attributeType === AttributeType.MultiselectPickList ||
    attributeType === AttributeType.TwoOptions;

  const needToGetFormattedValue2 = () =>
    attributeType === AttributeType.LookUp ||
    attributeType === AttributeType.Owner ||
    attributeType === AttributeType.Customer ||
    attributeType === AttributeType.TwoOptions;

  const needToGetFormattedValue3 = () =>
    attributeType === AttributeType.LookUp ||
    attributeType === AttributeType.Owner ||
    attributeType === AttributeType.Customer;

  if (attributeType === AttributeType.WholeNumber) {
    const format: string = entityMetadata.Attributes._collection[fieldName].Format;
    const field: string = getWholeNumberFieldName(format, entity, fieldName, timeZoneDefinitions);
    return [field, false];
  }
  else if (needToGetFormattedValue1()) {
    return [entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`], false];
  }
  else if (isLinkEntity && needToGetFormattedValue2()) {
    return [entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`], true];
  }
  else if (fieldName === entityMetadata._primaryNameAttribute) {
    return [entity[fieldName], true];
  }
  else if (needToGetFormattedValue3()) {
    return [entity[`_${fieldName}_value@OData.Community.Display.V1.FormattedValue`], true];
  }
  return [entity[fieldName], false];
};

const genereateItems = (props: IItemProps): Entity => {
  const {
    item,
    isLinkEntity,
    attributeType,
    fieldName,
    entity,
    fetchXml,
    index,
  } = props;

  const hasAggregate: boolean = isAggregate(fetchXml ?? '');
  const entityName: string = getEntityName(fetchXml ?? '');

  if (hasAggregate) {
    const aggregateAttrNames: string[] = getAliasNames(fetchXml ?? '');

    return item[aggregateAttrNames[index]] = {
      displayName: entity[aggregateAttrNames[index]],
      isLinkable: false,
      entity,
      fieldName: aggregateAttrNames[index],
      attributeType,
      entityName,
      isLinkEntity,
      aggregate: true,
    };
  }

  const [displayName, isLinkable] = getEntityData(props);

  return item[fieldName] = {
    displayName,
    isLinkable,
    entity,
    fieldName,
    attributeType,
    entityName,
    isLinkEntity,
  };
};

export const getTopInFetchXml = (fetchXml: string | null): number => {
  if (!fetchXml) return 0;
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');
  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

  const top: string | null = fetch?.getAttribute('top');

  return Number(top) || 0;
};

export const getItems = async (
  fetchXml: string | null,
  pageSize: number,
  currentPage: number,
  recordsCount: number): Promise<Entity[]> => {

  const pagingFetchData: string = addPagingToFetchXml(
    fetchXml ?? '',
    pageSize,
    currentPage,
    recordsCount);

  const attributesFieldNames: string[] = getAttributesFieldNames(pagingFetchData);
  const entityName: string = getEntityName(fetchXml ?? '');
  const records: RetriveRecords = await getCurrentPageRecords(pagingFetchData);

  const entityMetadata: EntityMetadata = await getEntityMetadata(entityName, attributesFieldNames);
  const linkEntityAttFieldNames: Dictionary<EntityAttribute[]> = getLinkEntitiesNames(
    fetchXml ?? '');

  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes: EntityAttribute[][] = Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => {
    const attributeNames: string[] = linkEntityAttributes[i].map(attr => attr.name);
    return getEntityMetadata(linkEntityNames, attributeNames);
  });

  const linkentityMetadata: EntityMetadata[] = await Promise.all(promises);

  const items: Entity[] = [];
  const timeZoneDefinitions = await getTimeZoneDefinitions();

  records.entities.forEach(entity => {
    const item: Entity = { id: entity[`${entityName}id`] };

    attributesFieldNames.forEach((fieldName, index) => {
      const attributeType: number = entityMetadata.Attributes.get(fieldName).AttributeType;

      const attributes: IItemProps = {
        timeZoneDefinitions,
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

    linkEntityNames.forEach((linkEntityName, i) => {
      linkEntityAttributes[i].forEach((attr, index) => {
        let fieldName = attr.attributeAlias;

        if (!fieldName && attr.linkEntityAlias) {
          fieldName = `${attr.linkEntityAlias}.${attr.name}`;
        }
        else if (!fieldName) {
          fieldName = `${linkEntityName}${i + 1}.${attr.name}`;
        }

        const attributeType: number = linkentityMetadata[i].Attributes.get(attr.name).AttributeType;

        const attributes: IItemProps = {
          timeZoneDefinitions,
          item,
          isLinkEntity: true,
          entityMetadata: linkentityMetadata[i],
          attributeType,
          fieldName,
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

export const setFilteredColumns = async (
  fetchXml: string | null,
  setColumns: React.Dispatch<React.SetStateAction<IColumn[]>>,
  getColumns: (fetchXml: string | null)=> Promise<IColumn[]>): Promise<void> => {

  const columns: IColumn[] = await getColumns(fetchXml);
  const order = getOrderInFetch(fetchXml ?? '');

  if (order) {
    const filteredColumns = columns.map(col => {
      if (col.ariaLabel === Object.values(order)[0] ||
      col.fieldName === Object.values(order)[0]) {

        col.isSorted = true;
        col.isSortedDescending = Object.keys(order)[0] === 'true';
        return col;
      }
      return col;
    });
    setColumns(filteredColumns);
  }
  else {
    setColumns(columns);
  }
};

export const getFilteredRecords = async (
  totalRecordsCount: React.MutableRefObject<number>,
  fetchXml: string | null,
  pageSize: number,
  currentPage: number,
  nextButtonDisabled: React.MutableRefObject<boolean>,
  lastItemIndex: React.MutableRefObject<number>,
  firstItemIndex: React.MutableRefObject<number>): Promise<Entity[]> => {
  const recordsCount: number = await getRecordsCount(fetchXml ?? '');

  totalRecordsCount.current = recordsCount;
  nextButtonDisabled.current = Math.ceil(recordsCount / pageSize) <= currentPage;

  const records: Entity[] = await getItems(
    fetchXml,
    pageSize,
    currentPage,
    recordsCount);

  lastItemIndex.current = (currentPage - 1) * pageSize + records.length;
  firstItemIndex.current = (currentPage - 1) * pageSize + 1;

  return records;
};
