import { AttributeType } from './enums';

export const checkIfAttributeIsEntityReferance = (attributeType: AttributeType): boolean =>
  attributeType === AttributeType.Lookup ||
  attributeType === AttributeType.Owner ||
  attributeType === AttributeType.Customer;

export const needToGetFormattedValue = (attributeType: AttributeType) =>
  attributeType === AttributeType.Money ||
  attributeType === AttributeType.PickList ||
  attributeType === AttributeType.DateTime ||
  attributeType === AttributeType.MultiselectPickList ||
  attributeType === AttributeType.TwoOptions;
