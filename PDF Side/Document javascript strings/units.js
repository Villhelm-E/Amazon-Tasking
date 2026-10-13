// set up variables
generated = false;
pathvar = this.path.replace(this.documentFileName,"");
attachment_name = "equipment_list";
full_path = pathvar + attachment_name + ".json";

// Refresh JSON file attachment
function reload_equipment_list() {
    if (this.dataObjects != null) {
        //Delete the existing equipment list
        this.removeDataObject(attachment_name);
    }

    //Import the new equipment list
    try {
        this.importDataObject(attachment_name, full_path);
    } catch(e) {
        app.alert("No equipment list found. \n\nMake sure to add \'" + pathvar + "\' to Trusted Sites", 0, 0, "Equipment List Not Found");
    }
}

// Read JSON attachment contents to "units" variable
function load_units() {
    
    // read from JSON data object
    try {
        var rStream = this.getDataObjectContents(attachment_name);
    } catch(e) {
        app.alert("Can't read equipment list. Make sure to disable \'Protected Mode at Startup\' in Security (Advanced) preference.", 0, 0, "Can't Read Equipment List");
    }

    console.println("created rStream");

    // save data object contents to string
    var cFile = util.stringFromStream(rStream);

    console.println("created cFile");
    
    // parse json contents
    try {
        units = JSON.parse(cFile);
        console.println("created units");
    } catch(e) {
        app.alert("No Equipment list found. Make sure to disable \'Protected Mode at Startup\' in Security (Advanced) preference.", 0, 0, "Equipment List Not Found");
    }
}

// Check if tasking has been generated
if ( this.getField("reload").hidden == false && this.getField("Submit").hidden == false ) {
    generated = false;
} else {
    generated = true;
}
// load content from json to "units" variable on document open
// but only if the tasking hasn't been generated
if (generated == false) {
    reload_equipment_list();
    load_units();
}