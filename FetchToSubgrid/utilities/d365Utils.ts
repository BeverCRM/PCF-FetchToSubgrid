import * as React from 'react';
import { IColumn } from '@fluentui/react';
import {
  getCurrentPageRecords,
  getEntityMetadata,
  getRecordsCount,
  getTimeZoneDefinitions,
  getWholeNumberFieldName,
} from '../services/dataverseService';
import { AttributeType } from './enums';
import {
  addPagingToFetchXml,
  getAliasNames,
  getAttributesFieldNames,
  getEntityName,
  getLinkEntitiesNames,
  getOrderInFetch,
  isAggregate } from './fetchXmlUtils';
import {
  Dictionary,
  Entity, EntityAttribute,
  EntityMetadata,
  IItemProps,
  RetriveRecords } from './types';
import {
  needToGetFormattedValue,
  checkIfAttributeIsEntityReferance,
} from './utils';

export const getEntityData = (props: IItemProps) => {
  const {
    timeZoneDefinitions,
    isLinkEntity,
    entityMetadata,
    attributeType,
    fieldName,
    entity,
  } = props;

  if (attributeType === AttributeType.Number) {
    const format: string = entityMetadata.Attributes._collection[fieldName].Format;
    const field: string = getWholeNumberFieldName(format, entity, fieldName, timeZoneDefinitions);
    return [field, false];
  }
  else if (needToGetFormattedValue(attributeType)) {
    return [entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`], false];
  }
  else if (isLinkEntity && checkIfAttributeIsEntityReferance(attributeType)) {
    return [entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`], true];
  }
  else if (fieldName === entityMetadata._primaryNameAttribute) {

    return [entity[fieldName], true];
  }
  else if (checkIfAttributeIsEntityReferance(attributeType)) {
    return [entity[`_${fieldName}_value@OData.Community.Display.V1.FormattedValue`], true];
  }
  return [entity[fieldName], false];
};

export const genereateItems = (props: IItemProps): Entity => {
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
  getColumns: (fetchXml: string | null) => Promise<IColumn[]>): Promise<void> => {

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
