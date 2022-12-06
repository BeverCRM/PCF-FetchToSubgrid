import { IInputs } from '../generated/ManifestTypes';
import utilities from '../utilities';

let _context: ComponentFramework.Context<IInputs>;

export default {
  setContext(context: ComponentFramework.Context<IInputs>) {
    _context = context;
  },

  getRecordsPerPage(): number {
    // @ts-ignore
    return _context.userSettings.pagingLimit;
  },

  async getRecordsCount(fetchXml: string): Promise<number> {
    const entityName = utilities.getEntityName(fetchXml);
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;
    const records = await _context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);
    return records.entities.length;
  },

  async getEntityMetadata(entityName: string, attributesFieldNames: string[]):
  Promise<ComponentFramework.PropertyHelper.EntityMetadata> {
    const entityMetadata = await _context.utils.getEntityMetadata(entityName, attributesFieldNames);
    return entityMetadata;
  },

  async getColumns(fetchXml: string | null):
   Promise<ComponentFramework.PropertyHelper.EntityMetadata> {

    const attributesFieldNames: string[] = utilities.getAttributesFieldNames(fetchXml ?? '');
    const entityName: string = utilities.getEntityName(fetchXml ?? '');

    const entityMetadata = await this.getEntityMetadata(entityName, attributesFieldNames);
    const displayNameCollection = entityMetadata.Attributes._collection;

    const linkEntityNameAndAttributes = utilities.getLinkEntitiesNames(fetchXml ?? '');
    const columns: Array<object> = [];

    const entityNames = Object.keys(linkEntityNameAndAttributes);
    const entityFieldNames: any = Object.values(linkEntityNameAndAttributes);
    const data: any = {};

    for (let i = 0; i < entityFieldNames.length; i++) {
      const onlyArrtibuteNames: string[] = [];

      for (let j = 1; j < entityFieldNames[i].length; j++) {
        onlyArrtibuteNames.push(entityFieldNames[i][j]);
      }
      // @ts-ignore
      const medData = await this.getEntityMetadata(entityNames[i], onlyArrtibuteNames);
      const DisplayNameCollection = medData.Attributes._collection;

      data[entityNames[i]] = DisplayNameCollection;
    }

    entityFieldNames.forEach((attr: any, index: number) => {
      const entityName: string = entityNames[index];

      for (let i = 1; i < attr.length; i++) {
        const alias: string = attr[0];
        const attributeName: string = attr[i];
        const display: string = data[entityName][attributeName].DisplayName;
        columns.push({
          name: `${display} (${alias})`,
          fieldName: `${alias}.${attributeName}`,
        });
      }
    });

    attributesFieldNames.forEach((name: string) => {
      columns.push({
        name: displayNameCollection[name].DisplayName,
        fieldName: name,
      });
    });

    return columns;
  },

  openLookupForm(entity: any, fieldName: string): void {
    _context.navigation.openForm(
      {
        entityName: entity[`_${fieldName}_value@Microsoft.Dynamics.CRM.lookuplogicalname`],
        entityId: entity[`_${fieldName}_value`],
      },
    );
  },

  openLinkEntityRecord(entity: any, fieldName: string) {
    const entityName = entity[`${fieldName}@Microsoft.Dynamics.CRM.lookuplogicalname`];
    _context.navigation.openForm(
      {
        entityName,
        entityId: entity[fieldName],
      },
    );
  },

  openPrimaryEntityForm(entity: any, entityName: string): void {
    _context.navigation.openForm(
      {
        entityName,
        entityId: entity[`${entityName}id`],
      },
    );
  },

  openRecord(entityName: string, entityId: string) {
    _context.navigation.openForm(
      {
        entityName,
        entityId,
      },
    );
  },

  async getRecords(fetchXml: string | null, entityName: string):
  Promise<ComponentFramework.WebApi.RetrieveMultipleResponse> {
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;

    return await _context.webAPI.retrieveMultipleRecords(entityName, encodeFetchXml);
  },
};
