import { IColumn } from '@fluentui/react';
import { AttributeType } from '../@types/enums';
import { needToGetFormattedValue, checkIfAttributeIsEntityReferance } from './utils';
import {
  addPagingToFetchXml,
  getEntityAggregateAliasNames,
  getAttributesFieldNamesFromFetchXml,
  getEntityNameFromFetchXml,
  getLinkEntitiesNamesFromFetchXml,
  isAggregate,
  getLinkEntityAggregateAliasNames,
  changeAliasNames,
} from './fetchXmlUtils';
import {
  Dictionary,
  Entity, EntityAttribute,
  EntityMetadata,
  IDataverseService,
  IItemProps,
  RetriveRecords,
} from '../@types/types';

const getLinkEntityFieldName = (
  changedAliasNames: string[],
  attr: EntityAttribute,
  index: number,
  linkEntityIndex: number,
  linkEntityName: string,
): string => {
  let fieldName = attr.attributeAlias;

  if (changedAliasNames[index]) {
    fieldName = changedAliasNames[index];
  }
  else if (!fieldName && attr.linkEntityAlias) {
    fieldName = `${attr.linkEntityAlias}.${attr.name}`;
  }
  else if (!fieldName) {
    fieldName = `${linkEntityName}${linkEntityIndex + 1}.${attr.name}`;
  }
  return fieldName;
};

const createColumnsForLinkEntity = (
  linkEntityNames: string[],
  linkEntityAttributes: EntityAttribute[][],
  linkentityMetadata: EntityMetadata[],
  fetchXml: string | null,
  columnWidth: number,
): IColumn[] => {
  const hasAggregate = isAggregate(fetchXml ?? '');
  const columns: IColumn[] = [];
  linkEntityNames.forEach((linkEntityName, i) => {
    linkEntityAttributes[i].forEach((attr, index) => {
      const changeAliasNameInFetch = changeAliasNames(fetchXml ?? '');
      const changedAliasNames: string[] | null =
        getLinkEntityAggregateAliasNames(changeAliasNameInFetch ?? '', i);

      const fieldName: string = getLinkEntityFieldName(
        changedAliasNames, attr, index, i, linkEntityName);

      const columnName: string = attr.attributeAlias ||
        linkentityMetadata[i].Attributes._collection[attr.name].DisplayName;
      const attributeType = linkentityMetadata[i].Attributes._collection[attr.name].AttributeType;

      const isMultiselectPickList = attributeType === AttributeType.MultiselectPickList;
      const sortingIsAllowed = isMultiselectPickList || hasAggregate || changedAliasNames[index];

      columns.push({
        styles: sortingIsAllowed
          ? { root: { '&:hover': { cursor: 'default' } } }
          : { root: { '&:hover': { cursor: 'pointer' } } },
        className: sortingIsAllowed
          ? 'colIsNotSortable'
          : 'linkEntity',
        ariaLabel: attr.name,
        name: columnName,
        fieldName,
        key: `col-el-${index}`,
        minWidth: 10,
        isResizable: true,
        isMultiline: false,
        calculatedWidth: columnWidth,
      });
    });
  });

  return columns;
};

const createColumnsForEntity = (
  attributesFieldNames: string[],
  displayNameCollection: Dictionary<EntityMetadata>,
  fetchXml: string | null,
  entityName: string,
  columnWidth: number,
): IColumn[] => {
  const columns: IColumn[] = [];

  attributesFieldNames.forEach((name, index) => {
    let displayName = name === `${entityName}id`
      ? 'Primary Key'
      : displayNameCollection[name].DisplayName;

    const attributeType: number = displayNameCollection[name].AttributeType;
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    const aliasNames: string[] | null = getEntityAggregateAliasNames(fetchXml ?? '');
    const isMultiselectPickList = attributeType === AttributeType.MultiselectPickList;

    const changeAliasNameInFetch = changeAliasNames(fetchXml ?? '');
    const changedAliasNames: string[] | null = getEntityAggregateAliasNames(changeAliasNameInFetch);
    const sortingIsAllowed = isMultiselectPickList || hasAggregate || changedAliasNames[index];

    if (aliasNames[index] && changedAliasNames?.length) {
      displayName = aliasNames[index];
      name = changedAliasNames[index];
    }

    columns.push({
      styles: sortingIsAllowed
        ? { root: { '&:hover': { cursor: 'default' } } }
        : { root: { '&:hover': { cursor: 'pointer' } } },
      className: sortingIsAllowed
        ? 'colIsNotSortable'
        : 'entity',
      name: displayName,
      fieldName: name,
      key: `col-${index}`,
      minWidth: 10,
      isResizable: true,
      isMultiline: false,
      calculatedWidth: columnWidth,
    });
  });

  return columns;
};

export const getEntityData = (props: IItemProps, dataverseService: IDataverseService) => {
  const {
    timeZoneDefinitions,
    isLinkEntity,
    entityMetadata,
    attributeType,
    fieldName,
    entity,
    hasAliasValue,
  } = props;

  if (attributeType === AttributeType.Number) {
    const format: string = entityMetadata.Attributes._collection[fieldName].Format;
    const field: string = dataverseService.getWholeNumberFieldName(
      format,
      entity,
      fieldName,
      timeZoneDefinitions);

    return [field, false];
  }

  const aliasAttrName = entity[`${fieldName}@OData.Community.Display.V1.AttributeName`];

  if (needToGetFormattedValue(attributeType)) {
    return [entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`], false];
  }

  if (isLinkEntity && checkIfAttributeIsEntityReferance(attributeType)) {
    return [entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`], true];
  }

  if (fieldName === entityMetadata._primaryNameAttribute ||
    aliasAttrName === entityMetadata._primaryNameAttribute) {
    return [entity[fieldName], true];
  }

  if (checkIfAttributeIsEntityReferance(attributeType)) {
    if (hasAliasValue) {
      return [entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`], true];
    }
    return [entity[`_${fieldName}_value@OData.Community.Display.V1.FormattedValue`], true];
  }
  return [entity[fieldName], false];
};

export const genereateItems = (props: IItemProps, dataverseService: IDataverseService): Entity => {
  const {
    item,
    isLinkEntity,
    attributeType,
    fieldName,
    entity,
    pagingFetchData,
    index,
    hasAliasValue,
  } = props;

  const hasAggregate: boolean = isAggregate(pagingFetchData ?? '');
  const entityName: string = getEntityNameFromFetchXml(pagingFetchData ?? '');

  if (hasAggregate) {
    const entityAggregateAttrNames: string[] = isLinkEntity
      ? getLinkEntityAggregateAliasNames(pagingFetchData ?? '', index)
      : getEntityAggregateAliasNames(pagingFetchData ?? '');

    return item[entityAggregateAttrNames[index]] = {
      displayName: entity[entityAggregateAttrNames[index]],
      isLinkable: false,
      entity,
      fieldName,
      attributeType,
      entityName,
      isLinkEntity,
      aggregate: true,
      hasAliasValue,
    };
  }

  const [displayName, isLinkable] = getEntityData(props, dataverseService);

  return item[fieldName] = {
    displayName,
    isLinkable,
    entity,
    fieldName,
    attributeType,
    entityName,
    isLinkEntity,
    hasAliasValue,
  };
};

const getRecordsData = async (
  fetchXml: string | null,
  pagingData: any,
  dataverseService: IDataverseService,
) => {
  const pagingFetchData: string = addPagingToFetchXml(
    fetchXml ?? '',
    pagingData.pageSize,
    pagingData.currentPage,
    pagingData.recordsCount);

  const attributesFieldNames: string[] = getAttributesFieldNamesFromFetchXml(pagingFetchData);
  const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
  const records: RetriveRecords = await dataverseService.getCurrentPageRecords(pagingFetchData);

  const entityMetadata: EntityMetadata = await dataverseService.getEntityMetadata(
    entityName, attributesFieldNames);
  const linkEntityAttFieldNames: Dictionary<EntityAttribute[]> = getLinkEntitiesNamesFromFetchXml(
    pagingFetchData ?? '');

  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes: EntityAttribute[][] = Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => {
    const attributeNames: string[] = linkEntityAttributes[i].map(attr => attr.name);
    return dataverseService.getEntityMetadata(linkEntityNames, attributeNames);
  });

  const linkentityMetadata: EntityMetadata[] = await Promise.all(promises);
  const entityAliases: string[] = getEntityAggregateAliasNames(pagingFetchData);
  const timeZoneDefinitions: Object = await dataverseService.getTimeZoneDefinitions();

  return {
    pagingFetchData,
    attributesFieldNames,
    entityName,
    records,
    entityMetadata,
    linkEntityNames,
    linkEntityAttributes,
    linkentityMetadata,
    timeZoneDefinitions,
    entityAliases,
  };
};

export const getItems = async (
  fetchXml: string | null,
  pageSize: number,
  currentPage: number,
  recordsCount: number,
  dataverseService: IDataverseService): Promise<Entity[]> => {
  const items: Entity[] = [];

  const {
    pagingFetchData,
    attributesFieldNames,
    entityName,
    records,
    entityMetadata,
    linkEntityNames,
    linkEntityAttributes,
    linkentityMetadata,
    timeZoneDefinitions,
    entityAliases,
  } = await getRecordsData(fetchXml, { pageSize, currentPage, recordsCount }, dataverseService);

  records.entities.forEach(entity => {
    const item: Entity = isAggregate(fetchXml ?? '') ? {} : { id: entity[`${entityName}id`] };
    attributesFieldNames.forEach((fieldName, index) => {
      const hasAliasValue = !!entityAliases[index];
      const attributeType: number = entityMetadata.Attributes.get(fieldName).AttributeType;

      if (entityAliases[index]) {
        fieldName = entityAliases[index];
      }

      const attributes: IItemProps = {
        timeZoneDefinitions,
        item,
        isLinkEntity: false,
        entityMetadata,
        attributeType,
        fieldName,
        entity,
        pagingFetchData,
        index,
        hasAliasValue,
      };

      genereateItems(attributes, dataverseService);
    });

    linkEntityNames.forEach((linkEntityName, i) => {
      linkEntityAttributes[i].forEach((attr, index) => {
        const hasAliasValue = !!attr.linkEntityAlias;
        const attributeType: number = linkentityMetadata[i].Attributes.get(attr.name).AttributeType;
        let fieldName = attr.attributeAlias;

        if (!fieldName && attr.linkEntityAlias) {
          fieldName = `${attr.linkEntityAlias}.${attr.name}`;
        }
        else if (!fieldName) {
          fieldName = `${linkEntityName}${i + 1}.${attr.name}`;
        }

        const attributes: IItemProps = {
          timeZoneDefinitions,
          item,
          isLinkEntity: true,
          entityMetadata: linkentityMetadata[i],
          attributeType,
          fieldName,
          entity,
          pagingFetchData,
          index,
          hasAliasValue,
        };

        genereateItems(attributes, dataverseService);
      });
    });

    items.push(item);
  });

  return items;
};

export const getColumns = async (
  fetchXml: string | null,
  allocatedWidth: number,
  dataverseService: IDataverseService): Promise<IColumn[]> => {
  const attributesFieldNames: string[] = getAttributesFieldNamesFromFetchXml(fetchXml ?? '');
  const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
  const entityMetadata: EntityMetadata = await dataverseService.getEntityMetadata(
    entityName,
    attributesFieldNames);

  const displayNameCollection: Dictionary<EntityMetadata> = entityMetadata.Attributes._collection;

  const linkEntityAttFieldNames: Dictionary<EntityAttribute[]> = getLinkEntitiesNamesFromFetchXml(
    fetchXml ?? '');
  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes: EntityAttribute[][] = Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => {
    const attributeNames: string[] = linkEntityAttributes[i].map(attr => attr.name);
    return dataverseService.getEntityMetadata(linkEntityNames, attributeNames);
  });
  const linkentityMetadata: EntityMetadata[] = await Promise.all(promises);

  const columnWidth = (allocatedWidth - 90) / attributesFieldNames.concat(linkEntityNames).length;

  const entityColumns = createColumnsForEntity(
    attributesFieldNames,
    displayNameCollection,
    fetchXml,
    entityName,
    columnWidth,
  );

  const linkEntityColumns = createColumnsForLinkEntity(
    linkEntityNames,
    linkEntityAttributes,
    linkentityMetadata,
    fetchXml,
    columnWidth,
  );

  return entityColumns.concat(linkEntityColumns);
};
