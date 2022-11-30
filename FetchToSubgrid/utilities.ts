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
