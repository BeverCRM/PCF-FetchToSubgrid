# Fetch to Subgrid

This control converts a FetchXml string into a subgrid.

<br>

![image](https://user-images.githubusercontent.com/60586462/233370479-48bf93e7-022a-4362-bb9f-99dda8452359.png)

Control has the following properties:

| Name | Type | Required | Description | Notes |
| ------------- | ------------- | ------------- | ------------- | ------------- |
| Default FetchXml | String | Required | The default FetchXml string to use when the *FetchXml Property* is blank. | The value must be a valid FetchXml string |
| Default Page Size | Number | Required | The default number of lines per page. |  |
| New Button Visibility | Boolean | Required | Enable the *New* button on the grid to create a new record. |  |
| Delete Button Visibility | Boolean | Required | Enable the *Delete* button on the grid to delete the selected records. |  |

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
