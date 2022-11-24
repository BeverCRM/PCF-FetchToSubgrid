import { IInputs } from '../generated/ManifestTypes';

let _context: ComponentFramework.Context<IInputs>;
let entityName: string;

export default {
  setContext(context: ComponentFramework.Context<IInputs>) {
    _context = context;
  },

  getAttributesFieldNames(fetchXml: string) {
    const parser = new DOMParser();
    const xmlDoc: any = parser.parseFromString(fetchXml, 'text/xml');
    const attributesFieldNames: string[] = [];
    entityName = xmlDoc.all[0].childNodes[1].attributes['name'].nodeValue;

    for (let i = 0; i < xmlDoc.all.length; i++) {
      if (xmlDoc.all[i].tagName === 'attribute') {
        attributesFieldNames.push(xmlDoc.all[i].attributes[0].value);
      }
    }
    return attributesFieldNames;
  },

  async getColumns(defaultFetchXml: string | null) {
    let fetchXml: string;

    if (_context.parameters.fetchXmlProperty.raw) {
      fetchXml = `${_context.parameters.fetchXmlProperty.raw}`;
    }
    else {
      fetchXml = `${defaultFetchXml}`;
    }

    const attributesFieldNames: string[] = this.getAttributesFieldNames(fetchXml);
    const entityMetadata = await _context.utils.getEntityMetadata(entityName, attributesFieldNames);

    const displayNameCollection = entityMetadata.Attributes._collection;
    const columns: any = [];

    attributesFieldNames.forEach((name: string, index: number) => {
      columns.push({
        name: displayNameCollection[name].DisplayName,
        fieldName: name,
        key: index,
      });
    });

    return columns;
  },

  async getItems(defaultFetchXml: string | null) {
    let fetchXml: string | null;

    if (_context.parameters.fetchXmlProperty.raw) {
      fetchXml = `${_context.parameters.fetchXmlProperty.raw}`;
    }
    else {
      fetchXml = `${defaultFetchXml}`;
    }

    const attributesFieldNames: string[] = this.getAttributesFieldNames(fetchXml);

    fetchXml = `?fetchXml=${encodeURIComponent(fetchXml)}`;

    const recors = await _context.webAPI.retrieveMultipleRecords(`${entityName}`, fetchXml);
    const items: any = [];

    recors.entities.forEach(entity => {
      const item: any = {};

      attributesFieldNames.forEach(fieldName => {
        if (fieldName in entity) {
          item[fieldName] = entity[fieldName];
        }

        if (`_${fieldName}_value` in entity) {
          item[fieldName] = entity[`_${fieldName}_value@OData.Community.Display.V1.FormattedValue`];
        }

      });

      items.push(item);
    });

    return items;
  },
};
