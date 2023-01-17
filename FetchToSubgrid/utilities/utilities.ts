/* global HTMLCollectionOf, NodeListOf */
import {
  getRecords,
  getEntityMetadata,
  getWholeNumberFieldName,
  getTimeZoneDefinitions } from '../services/crmService';
import { AttributeType } from './enums';

type Entity = ComponentFramework.WebApi.Entity;
type EntityMetadata = ComponentFramework.PropertyHelper.EntityMetadata;
type IItemProps = {
  timeZoneDefinitions: any;
  item: Entity;
  isLinkEntity: boolean;
  entityMetadata: EntityMetadata;
  attributeType: number;
  fieldName: string;
  entity: Entity;
  fetchXml: string | null;
  index: number;
}
interface EntityAttribute {
  linkEntityAlias: string | undefined;
  name: string;
  attributeAlias: string;
}

export const addPagingToFetchXml =
 (fetchXml: string, pageSize: number, currentPage: number): string => {
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

export const getLinkEntitiesNames = (fetchXml: string): { [key: string]: EntityAttribute[] } => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const linkEntities: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('link-entity');
  const linkEntityData: { [key: string]: EntityAttribute[] } = {};

  Array.prototype.slice.call(linkEntities).forEach(linkentity => {
    const entityName: string = linkentity.attributes['name'].value;
    const entityAttributes: EntityAttribute[] = [];

    const attributesSelector = `link-entity[name="${entityName}"] > attribute`;
    const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributesSelector);
    const linkEntityAlias: string | undefined =
     linkentity.attributes['alias'] && linkentity.attributes['alias'].value;

    Array.prototype.slice.call(attributes).map(attr => {
      const attributeAlias: string = attr.attributes['alias'] ? attr.attributes['alias'].value : '';

      entityAttributes.push(
        {
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

const genereateItems = (props: IItemProps): Entity => {
  const {
    timeZoneDefinitions,
    item,
    isLinkEntity,
    entityMetadata,
    attributeType,
    fieldName,
    entity,
    fetchXml,
    index } = props;

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

  if (attributeType === AttributeType.WHOLE_NUMBER) {
    const format: string = entityMetadata.Attributes._collection[fieldName].Format;
    const field: string = getWholeNumberFieldName(format, entity, fieldName, timeZoneDefinitions);
    displayName = field;
  }
  else if (attributeType === AttributeType.MONEY ||
      attributeType === AttributeType.PICKLIST ||
      attributeType === AttributeType.DATE_TIME ||
      attributeType === AttributeType.MULTISELECT_PICKLIST ||
      attributeType === AttributeType.TWO_OPTIONS) {

    displayName = entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`];
  }
  else if (isLinkEntity) {
    if (attributeType === AttributeType.LOOKUP ||
        attributeType === AttributeType.OWNER ||
        attributeType === AttributeType.CUSTOMER ||
        attributeType === AttributeType.TWO_OPTIONS) {
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
  else if (attributeType === AttributeType.LOOKUP ||
    attributeType === AttributeType.OWNER ||
    attributeType === AttributeType.CUSTOMER) {
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
  if (fetchXml) {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');
    const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

    const count: string | null = fetch.getAttribute('count');
    const top: string | null = fetch.getAttribute('top');

    return Number(count) || Number(top);
  }
  return 0;
};

export const getItems = async (
  fetchXml: string | null,
  pageSize: number,
  currentPage: number): Promise<Entity[]> => {

  const countOfRecordsInFetch = getCountInFetchXml(fetchXml);

  const pagingFetchData: string = addPagingToFetchXml(
    fetchXml ?? '', countOfRecordsInFetch || pageSize, currentPage);

  const attributesFieldNames: string[] = getAttributesFieldNames(pagingFetchData);
  const entityName: string = getEntityName(fetchXml ?? '');
  const records: ComponentFramework.WebApi.RetrieveMultipleResponse =
   await getRecords(pagingFetchData);

  const entityMetadata: EntityMetadata = await getEntityMetadata(entityName, attributesFieldNames);
  const linkEntityAttFieldNames: { [key: string]: EntityAttribute[] } =
   getLinkEntitiesNames(fetchXml ?? '');
  const linkEntityNames: string[] = Object.keys(linkEntityAttFieldNames);
  const linkEntityAttributes:Array<Array<EntityAttribute>> =
   Object.values(linkEntityAttFieldNames);

  const promises = linkEntityNames.map((linkEntityNames, i) => {
    const attributeNames: string[] = linkEntityAttributes[i].map(attr => attr.name);
    return getEntityMetadata(linkEntityNames, attributeNames);
  });

  const linkentityMetadata: EntityMetadata[] = await Promise.all(promises);
  const items: Entity[] = [];
  const timeZoneDefinitions = await getTimeZoneDefinitions();

  records.entities.forEach(entity => {
    const item: Entity = {
      id: entity[`${entityName}id`],
    };

    attributesFieldNames.forEach((fieldName, index) => {
      const attributeType: number = entityMetadata.Attributes._collection[fieldName].AttributeType;

      const attributes = {
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

        const attributeType: number =
      linkentityMetadata[i].Attributes._collection[attr.name].AttributeType;

        const attributes = {
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
