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
    // @ts-ignore
    const contextPage = _context.page;
    const clientUrl = contextPage.getClientUrl();
    const entityName = utilities.getEntityName(fetchXml);

    const result = await fetch(`${clientUrl}/api/data/v9.0/${entityName}s/?$count=true`);
    const records = await result.json();

    return records['@odata.count'];
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
    const columns: Array<object> = [];

    attributesFieldNames.forEach((name: string, index: number) => {
      columns.push({
        name: displayNameCollection[name].DisplayName,
        fieldName: name,
        key: index,
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

  openPrimaryEntityForm(entity: any, entityName:string): void {
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

  async getRecord(fetchXml: string | null, entityName: string):
  Promise<ComponentFramework.WebApi.RetrieveMultipleResponse> {
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;

    return await _context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);
  },
};
