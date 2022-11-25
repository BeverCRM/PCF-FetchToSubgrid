import { IInputs } from '../generated/ManifestTypes';
import utilities from '../utilities';

let _context: ComponentFramework.Context<IInputs>;

export default {
  setContext(context: ComponentFramework.Context<IInputs>) {
    _context = context;
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

  async getRecord(fetchXml: string | null, entityName: string):
  Promise<ComponentFramework.WebApi.RetrieveMultipleResponse> {
    const encodeFetchXml: string = `?fetchXml=${encodeURIComponent(fetchXml ?? '')}`;

    return await _context.webAPI.retrieveMultipleRecords(`${entityName}`, encodeFetchXml);
  },
};
