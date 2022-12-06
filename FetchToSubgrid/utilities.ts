export default {
  addPagingToFetchXml(fetchXml: string, pageSize: number, currentPage: number) {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
    const fetch = xmlDoc.getElementsByTagName('fetch');

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

    const linkEntities = xmlDoc.getElementsByTagName('link-entity');
    const linkEntityData: any = {};

    Array.prototype.slice.call(linkEntities).forEach(linkentity => {
      const entityName: string = linkentity.attributes['name'].value;
      const alias: string = linkentity.attributes['alias'].value;
      const entityAttributes: string[] = [alias];
      const attributes = xmlDoc.querySelectorAll(`link-entity[name="${entityName}"] > attribute`);

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

    // @ts-ignore
    const entityName: string =
    xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
    const attributesFieldNames: string[] = [];

    const attributes = xmlDoc.querySelectorAll(`entity[name="${entityName}"] > attribute`);

    Array.prototype.slice.call(attributes).map(attr => {
      attributesFieldNames.push(attr.attributes.name.value);
    });

    return attributesFieldNames;
  },
};
