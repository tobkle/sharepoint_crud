# Sharepoint CRUD Operations

Javascript library which provides Sharepoint List operations for create, read, update, delete. 
They can be used, if you are already authenticated. So if you are using Sharepoint as a Web Server for these files.

* site.list.createListItem(newItem) - Creates a new List Item
* site.list.readListItem(itemId) - Reads a list item
* site.list.readListItems(numberOfRows) - Reads numberOfRows list items (or max 9999 records)
* site.list.updateListItem(itemId, updateItem) - Updates an item by id
* site.list.deleteListItem(itemId) - Deletes a list item by id
* site.list.getListFields() - Get List Field Descriptions
