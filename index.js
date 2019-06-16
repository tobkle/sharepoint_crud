// Main Routine
window.onload = function () {
    // Url of the Sharepoint Site
    var SharepointSite = "https://sharepoint.com"; // adjust this
    // site contains the Lists of the Sharepoint Site after connect
    var site = null;
    // CRUD is the Name (Title) of the Sharepoint List for this Example
    var list = "CRUD";

    // Get Fields Details of the List
    function getFieldDetails(){
        site[list].getListFields()
            .then(function (response) {
                if (response) console.log(JSON.parse(response).d.results);
            })
            .catch(function (error) {
                if (error) console.log(error);
            });
    }

    // Create a List Item
    function createItem(item) {
        site[list].createListItem(item)
            .then(function (response) {
                if (response) console.log(JSON.parse(response).d);
            })
            .catch(function (error) {
                if (error) console.log(error);
            });
    }

    // Read multiple List Items
    function readItems() {
        site[list].readListItems()
            .then(function (response) {
                if (response) console.log(JSON.parse(response).d.results);
            })
            .catch(function (error) {
                if (error) console.log(error);
            });
    }

    // Read single specific List Item
    function readItem(itemId) {
        site[list].readListItem(itemId)
            .then(function (response) {
                if (response) console.log(JSON.parse(response).d);
            })
            .catch(function (error) {
                if (error) console.log(error);
            });
    }

    // Update a List Item
    function updateItem(itemId, updateItem) {
        site[list].updateListItem(itemId, updateItem)
            .then(function (response) {
                if (response) console.log(response);
            })
            .catch(function (error) {
                if (error) console.log(error);
            });
    }

    // Delete a List Item
    function deleteItem(itemId) {
        site[list].deleteListItem(itemId)
            .then(function (response) {
                if (response) console.log(response);
            })
            .catch(function (error) {
                if (error) console.log(error.message);
            });
    }

    // Do CRUD operations on a Sharepoint list
    function doSomething() {
        // Show all Sharepoint Site Lists
        console.log(site);
        getFieldDetails();
        createItem({ Title: "My new List Item" });
        readItems();
        readItem(1);
        updateItem(1, { Title: "My updated List Item" });
        deleteItem(1);
    }

    // SPCRUD is the global variable which contains the functions to access
    // the Sharepoint List see source code in "sharepoint_crud.js"
    // Provides the function "connect" to connect to a Sharepoint Site
    SPCRUD.connect(SharepointSite)
        .then(function(SPSite) {
            // now, we've received an object with all Lists of the Site
            site = SPSite;
            doSomething();
        }).catch(function (error) {
            // we weren't able to reach the Sharepoint Site
            console.log(error);
        });
};
