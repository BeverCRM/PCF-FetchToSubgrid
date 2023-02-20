import { AttributeType } from './enums';

export const checkIfAttributeIsEntityReferance = (attributeType: AttributeType): boolean =>
  attributeType === AttributeType.LookUp ||
  attributeType === AttributeType.Owner ||
  attributeType === AttributeType.Customer;

export const needToGetFormattedValue1 = (attributeType: AttributeType) =>
  attributeType === AttributeType.Money ||
  attributeType === AttributeType.PickList ||
  attributeType === AttributeType.DateTime ||
  attributeType === AttributeType.MultiselectPickList ||
  attributeType === AttributeType.TwoOptions;

export const needToGetFormattedValue2 = (attributeType: AttributeType) =>
  attributeType === AttributeType.LookUp ||
  attributeType === AttributeType.Owner ||
  attributeType === AttributeType.Customer ||
  attributeType === AttributeType.TwoOptions;

export const needToGetFormattedValue3 = (attributeType: AttributeType) =>
  attributeType === AttributeType.LookUp ||
  attributeType === AttributeType.Owner ||
  attributeType === AttributeType.Customer;
