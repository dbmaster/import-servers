Inventory-import plugin synchronizes objects from excel file with project data.
Current implementation supports importing Applications, Servers, and Connections.

## Parameters

| Parameter name         | Type   | Required | Description                                |
|------------------------|:------:|:--------:|--------------------------------------------|
| File to Import (Excel) | File   | Yes      | Data will be imported from this file. See format and sample below
| Object Type            | List   | Yes      | Type of object to be imported. Possible values are Server, Application, and Connection. |
| Inventory Object Filter| String | No       | Filter for the objects in the repository in [dbmaster search format](https://www.dbmaster.io/documentation/quick-search). This parameter is useful when importing from more than one source or when some objects are manually entered and some come from external data. |
| Data Source            | String | No       | When non-empty, custom field "Source" will be set to this value|
| Action                 | List   | Yes      | Allows previewing changes to be made before importing objects.  Possible values: <ul><li>Preview - output of the plugin will contain changes to be made to the data without any changes made</li><li>Import - data will be imported if it does not have any errors</li></ul>
| Field Mapping          | Text   | No       | Specifies mapping between columns in the source file and inventory fields. Mappings should be defined one line in format &lt;_excel-column_&gt;=&lt;_custom-field_&gt;. A column will be ignored when _custom-field_ is empty. If there is no mapping for a column, it will be imported under the same name. |

## Import File Format

The plugin imports data from the first sheet of excel file, others are simply ignored.
First row must have column names. 

For **date fields** the plugin recognizes these formats 

* MMM d, yyyy
* MMM d,yyyy h:mm a
* MMM d,yyyy h:mm:ss
* EEE MMM d, yyyy h:mm a
* EEE MMM d h:mm:ss z yyyy
 
For details and examples see [original documentation](http://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html)

## Import File Sample
