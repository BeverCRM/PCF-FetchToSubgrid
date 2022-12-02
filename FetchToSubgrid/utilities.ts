export default {

  getEntityName(fetchXml: string): string {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

    return xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
  },

  getAttributesFieldNames(fetchXml: string): string[] {
    const parser: DOMParser = new DOMParser();
    const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

    // eslint-disable-next-line no-undef
    const attributes: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('attribute');

    const attributesFieldNames: string[] =
    Array.prototype.slice.call(attributes).map(attr => attr.getAttribute('name'));

    return attributesFieldNames;
  },
};
