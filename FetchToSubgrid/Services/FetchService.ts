import { IInputs } from '../generated/ManifestTypes';

let _context: ComponentFramework.Context<IInputs>;

export default {
  setContext(context: ComponentFramework.Context<IInputs>) {
    _context = context;
  },

  findCurrentText(str: string): string {
    const startIndex: number = str.indexOf('<entity name');

    for (let i = startIndex; ; i++) {
      if (str[i] === '>') {
        return str.slice(startIndex + 14, i - 1);
      }
    }
  },

  async getFetchData() {
    const inputValue = _context.parameters.sampleProperty.raw;
    let fetchXml: string = `${inputValue}`;
    const entityName: string = this.findCurrentText(fetchXml);

    fetchXml = `?fetchXml=${encodeURIComponent(fetchXml)}`;
    const recordRelatedNotes = await _context.webAPI.retrieveMultipleRecords(`${entityName}`,
      fetchXml);

    return recordRelatedNotes.entities;
  },
};
