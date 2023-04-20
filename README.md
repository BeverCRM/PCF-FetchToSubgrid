# Fetch to Subgrid

This control converts a FetchXml string into a subgrid.

<br>

![image](https://user-images.githubusercontent.com/60586462/233363262-596dde89-6b21-4c52-a73c-1aafb653f834.png)

Control has the following properties:

| Name | Type | Required | Description | Notes |
| ------------- | ------------- | ------------- | ------------- | ------------- |
| FetchXml Property | Sting | Required | The *Multiple Lines of Text* field to which the control is bound. | The value must be a valid FetchXml string or a valid JSON with the following properties: <br><br>* newButtonVisibility (boolean, optional) <br>* deleteButtonVisibility (boolean, optional) <br>* pageSize (number, optional) <br>* fetchXml (string, optional) |
| Default FetchXml | String | Required | The default FetchXml string to use when the *FetchXml Property* is blank. | The value must be a valid FetchXml string |
| Default Page Size | Number | Required | The default number of lines per page. |  |
| New Button Visibility | Boolean | Required | Enable the *New* button on the grid to create a new record. |  |
| Delete Button Visibility | Boolean | Required | Enable the *Delete* button on the grid to delete the selected records. |  |

<br>

![image](https://user-images.githubusercontent.com/60586462/233362848-3acb9a0f-9478-4e54-8763-c84b98e93aa9.png)

If the FetchXml Property is a JSON string and doesn't contain any of the four properties listed in the table above, then the default values are used.

You can use almost the all possibilities of a FetchXml string:
* FetchXML Aggregation
* Multiple Link-Entities
* The fetch tag attributes
* etc.
