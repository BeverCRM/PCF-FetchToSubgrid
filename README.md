# Fetch to Subgrid

This control converts a _"Multiple Lines of Text"_ field into a subgrid.

<br>

![image](https://user-images.githubusercontent.com/60586462/233377327-4e43c785-96fa-4845-8308-d523ce457f58.png)

<br>

The _"Multiple Lines of Text"_ field value can contain either a FetchXml string or a JSON string with the following format:

```json
{
    "newButtonVisibility": true,
    "deleteButtonVisibility": true,
    "pageSize": 10,
    "fetchXml": "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">\n  <entity name=\"bvr_testcase\">\n    <attribute name=\"bvr_name\" />\n    <attribute name=\"bvr_slot_title\" />\n    <attribute name=\"bvr_os_status\" />\n    <attribute name=\"createdon\" />\n    <attribute name=\"statuscode\" />\n    <attribute name=\"statecode\" />\n    <attribute name=\"ownerid\" />\n    <filter type=\"and\">\n      <condition attribute=\"bvr_os_status\" operator=\"eq\" value=\"551800000\" />\n    </filter>\n  </entity>\n</fetch>"
}
```

<br>

| Property | Description |
| ------------- | ------------- |
| fetchXml | FetchXml query, according to which the data will be shown in the sub-grid. |
| pageSize | Number of sub-grid lines per page. |
| newButtonVisibility | Show or hide _New_ button on the sub-grid. |
| deleteButtonVisibility | Show or hide _Delete_ button on the sub-grid. |

<br>

Control has the following properties:

| Name | Type | Required | Description |
| ------------- | ------------- | ------------- | ------------- |
| Default FetchXml | String | Required | The default FetchXml value will be used if the _"Multiple Lines of Text"_ field value doesn't contain the _"fetchXml"_ JSON property. |
| Default Page Size | Number | Required | The default Page Size value will be used if the _"Multiple Lines of Text"_ field value doesn't contain the _"pageSize"_ JSON property. |
| New Button Visibility | Boolean | Required | The default New Button Visibility will be used if the _"Multiple Lines of Text"_ field value doesn't contain the _"newButtonVisibility"_ JSON property. |
| Delete Button Visibility | Boolean | Required | The default Delete Button Visibility will be used if the _"Multiple Lines of Text"_ field value doesn't contain the _"deleteButtonVisibility"_ JSON property. |

<br>

![image](https://user-images.githubusercontent.com/60586462/233362848-3acb9a0f-9478-4e54-8763-c84b98e93aa9.png)

<br>

FetchXml can contain:
* FetchXML aggregate columns
* Multiple link entities
* Attributes of the fetch tag
* etc.
