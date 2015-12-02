Inventory-import plugin synchronizes objects from excel file with project data.
Current implementation supports importing Applications, Servers, and Connections.

## Parameters

| Parameter name         | Type   | Required | Description                                |
|------------------------|:------:|:--------:|--------------------------------------------|
| File to Import (Excel) | File   | Yes      | Data will be imported from this file. See format and sample below
| Object Type            | List   | Yes      | Type of object to be imported. Possible values are Server, Application, Connection. |
| Inventory Object Filter| String | No       | Filter for the objects in the repository in [dbmaster search format](https://www.dbmaster.io/documentation/quick-search). This parameter is usefull when importing from more than one source or when some objects are manually entered and some come from external data. |
| Data Source            | String | No       | When non-empty, custom field "Source" will be set to this value|
| Action                 | List   | Yes      | Allows preview changes to be made before importing objects.  Possible values: <ul><li>Preview - output of the plugin will contain changes to be made to the data wihtout any changes made</li><li>Import - data will be imported if it does not have any errors</li></ul>
| Field Mapping          | Text   | No       | 

## Import File Format

### Date Format


## Import File Sample
