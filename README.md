# Fetch to Subgrid

This control converts a _"Multiple Lines of Text"_ field into a subgrid.

<br>

![image](https://user-images.githubusercontent.com/60586462/233377327-4e43c785-96fa-4845-8308-d523ce457f58.png)

The _"Multiple Lines of Text"_ field value can contain either a FetchXml string or a JSON string with the following format:

<br>

```json
{
    "newButtonVisibility": true,
    "deleteButtonVisibility": true,
    "pageSize": 10,
    "fetchXml": "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">\n  <entity name=\"bvr_testcase\">\n    <attribute name=\"bvr_name\" />\n    <attribute name=\"bvr_slot_title\" />\n    <attribute name=\"bvr_os_status\" />\n    <attribute name=\"createdon\" />\n    <attribute name=\"statuscode\" />\n    <attribute name=\"statecode\" />\n    <attribute name=\"ownerid\" />\n    <filter type=\"and\">\n      <condition attribute=\"bvr_os_status\" operator=\"eq\" value=\"551800000\" />\n    </filter>\n  </entity>\n</fetch>"
}
```

Control has the following properties:

| Name | Type | Required | Description |
| ------------- | ------------- | ------------- | ------------- |
| Default FetchXml | String | Required | The default FetchXml string to use when the *FetchXml Property* is blank. |
| Default Page Size | Number | Required | The default number of lines per page. |
| New Button Visibility | Boolean | Required | Enable the *New* button on the grid to create a new record. |
| Delete Button Visibility | Boolean | Required | Enable the *Delete* button on the grid to delete the selected records. |

<br>

![image](https://user-images.githubusercontent.com/60586462/233362848-3acb9a0f-9478-4e54-8763-c84b98e93aa9.png)

The **FetchXml Property** is the underlying *Multiple Lines of Text* field to which the control is bound.
The value must be a valid FetchXml string or a valid JSON with the following properties:
- newButtonVisibility (boolean | optional)
- deleteButtonVisibility (boolean | optional)
- pageSize (number | optional)
- fetchXml (string | optional)

<br>

FetchXml can contain:
* FetchXML aggregate columns
* Multiple link entities
* Attributes of the fetch tag
* etc.

<br>

> **Note** If the **FetchXml Property** is a JSON string and doesn't contain any of the four properties listed above, then the default values are used.
