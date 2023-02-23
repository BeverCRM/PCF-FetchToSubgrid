import { ColumnActionsMode, IColumn } from '@fluentui/react';
import { AttributeType } from './enums';
import { needToGetFormattedValue, checkIfAttributeIsEntityReferance } from './utils';
import {
  addPagingToFetchXml,
  getAliasNames,
  getAttributesFieldNamesFromFetchXml,
  getEntityNameFromFetchXml,
  getLinkEntitiesNamesFromFetchXml,
  isAggregate,
} from './fetchXmlUtils';
import {
  Dictionary,
  Entity, EntityAttribute,
  EntityMetadata,
  IDataverseService,
  IItemProps,
  RetriveRecords,
} from './types';

const createColumnsForLinkEntity = (
  linkEntityNames: string[],
  linkEntityAttributes: EntityAttribute[][],
  linkentityMetadata: EntityMetadata[],
): IColumn[] => {
  const columns: IColumn[] = [];
  linkEntityNames.forEach((linkEntityName, i) => {
    linkEntityAttributes[i].forEach((attr, index) => {
      let fieldName = attr.attributeAlias;

      if (!fieldName && attr.linkEntityAlias) {
        fieldName = `${attr.linkEntityAlias}.${attr.name}`;
      }
      else if (!fieldName) {
        fieldName = `${linkEntityName}${i + 1}.${attr.name}`;
      }

      const columnName: string = attr.attributeAlias ||
      linkentityMetadata[i].Attributes._collection[attr.name].DisplayName;

      columns.push({
        className: 'linkEntity',
        ariaLabel: attr.name,
        name: columnName,
        fieldName,
        key: `col-el-${index}`,
        minWidth: 10,
        isResizable: true,
        isMultiline: false,
        columnActionsMode: ColumnActionsMode.hasDropdown,
      });
    });
  });
  return columns;
};

const createColumnsForEntity = (
  attributesFieldNames: string[],
  displayNameCollection: Dictionary<EntityMetadata>,
  fetchXml: string | null,
): IColumn[] => {
  const columns: IColumn[] = [];
  attributesFieldNames.forEach((name, index) => {
    let displayName = displayNameCollection[name].DisplayName;
    const hasAggregate: boolean = isAggregate(fetchXml ?? '');
    const aggregateAttrNames: string[] | null = hasAggregate ? getAliasNames(fetchXml ?? '') : null;

    if (aggregateAttrNames?.length) {
      displayName = aggregateAttrNames[index];
      name = aggregateAttrNames[index];
    }

    columns.push({
      className: 'entity',
      name: displayName,
      fieldName: name,
      key: `col-${index}`,
      minWidth: 10,
      isResizable: true,
      isMultiline: false,
      columnActionsMode: ColumnActionsMode.hasDropdown,
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

export const genereateItems = (props: IItemProps, dataverseService: IDataverseService): Entity => {
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
  const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');

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

  const [displayName, isLinkable] = getEntityData(props, dataverseService);

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
  recordsCount: number,
  dataverseService: IDataverseService): Promise<Entity[]> => {

  const pagingFetchData: string = addPagingToFetchXml(
    fetchXml ?? '',
    pageSize,
    currentPage,
    recordsCount);

  const attributesFieldNames: string[] = getAttributesFieldNamesFromFetchXml(pagingFetchData);
  const entityName: string = getEntityNameFromFetchXml(fetchXml ?? '');
  const records: RetriveRecords = await dataverseService.getCurrentPageRecords(pagingFetchData);

  const entityMetadata: EntityMetadata = await dataverseService.getEntityMetadata(
    entityName, attributesFieldNames);
  const linkEntityAttFieldNames: Dictionary<EntityAttribute[]> = getLinkEntitiesNamesFromFetchXml(
    fetchXml ?? '');

  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes: EntityAttribute[][] = Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => {
    const attributeNames: string[] = linkEntityAttributes[i].map(attr => attr.name);
    return dataverseService.getEntityMetadata(linkEntityNames, attributeNames);
  });

  const linkentityMetadata: EntityMetadata[] = await Promise.all(promises);

  const items: Entity[] = [];
  const timeZoneDefinitions = await dataverseService.getTimeZoneDefinitions();

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

      genereateItems(attributes, dataverseService);
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

        genereateItems(attributes, dataverseService);
      });
    });

    items.push(item);
  });

  return items;
};

export const getColumns = async (
  fetchXml: string | null,
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

  const entityColumns = createColumnsForEntity(
    attributesFieldNames,
    displayNameCollection,
    fetchXml,
  );

  const linkEntityColumns = createColumnsForLinkEntity(
    linkEntityNames,
    linkEntityAttributes,
    linkentityMetadata,
  );

  return entityColumns.concat(linkEntityColumns);
};
