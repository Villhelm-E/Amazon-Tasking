function clean() {
    //----------------------------------------------------------
    // Load fields into an array
    //----------------------------------------------------------
    
    // Set up empty array
    var fieldArray = [];

    // loop through all fields in the document
    for (var i = 0; i < this.numFields - 1; i++) {
        
        //Get field name
        var fieldName = this.getNthFieldName(i);
        var field = this.getField(fieldName);

        // Skip if the field object is null (e.g., if the field was deleted)
        if (field == null) {
            continue;
        }

        // Skip if the field is on the first page or hidden
        var fieldPage = field.page;
        if ( ( Number.isInteger(fieldPage) == false && typeof fieldPage != "object" ) || fieldPage < 1) {
            continue;
        }

        // Grab the final part of the field
        var suffix = fieldName.substring(fieldName.lastIndexOf("."));

        //Identify UnitID field as "UnitID"
        if (suffix == ".UnitID") {
            var fieldType = "UnitID";
        } else {
            // Identify Header fields as "Header"
            if (suffix == ".UnitType" || suffix == ".SubType" || suffix == ".DriveType" || suffix == ".manufacturer" || suffix == ".model" || suffix == ".serial" || suffix == ".tech") {
                var fieldType = "Header";
            //Everything else is a Task field
            } else {
                var fieldType = "Task";
            }
        }

        // Handle null case for field.value
        if (field.value == null) {
            var fieldValue = "";
        } else {
            var fieldValue = field.value;
        }

        //Check if field appears once or multiple times (in the case of combining tasking sheets)
        if (typeof fieldPage == "object") {
            //If fieldPage is an array, loop through each and add to UnitPages[]
            for (j = 0; j < fieldPage.length - 1; j++) {
                // Add the field's name, value, page number, and type as key-pairs in UnitPages[]
                fieldArray.push( {'name':fieldName, 'val':fieldValue, 'page':fieldPage[j], 'type':fieldType} );
            }
        //To-Do: Rewrite this so the push action isn't repeated
        } else {
            // Add the field's name, value, page number, and type as key-pairs in UnitPages[]
            fieldArray.push( {'name':fieldName, 'val':fieldValue, 'page':field.page, 'type':fieldType} );
        }
    }

    
    //-------------------
    // Delete blank pages
    //-------------------

    // Sort fieldArray[] by page and store in sortedPages[]
    var sortedPages = fieldArray.sort(function(a, b){ return a.page - b.page; });

    // Start page count variable
    var pageCount = this.numPages; //This count will change
    var pC = pageCount; //This count will not change

    //Start the progress bar
    var t = app.thermometer;
    t.duration = pC;
    t.begin();
    var p = 0;
    
    // Loop through sortedPages[] in reverse order
    for (k = pageCount - 1; k > 0; k--) { // Subtract 1 from pageCount because 0-indexed pages
        
        //update thermometer stage
        t.text = "Processing page " + p + " of " + pC;
        p += 1;
        
        // Filter array to just page k
        var filterPage = sortedPages.filter((p) => p.page == k);

        // Filter array to just Task fields
        var tasks = filterPage.filter((f) => f.type == "Task");
        
        // Check if any fields are filled in
        if ( taskingDone(tasks) ) {
            // If tech filled in at least one Task field, skip page
            continue;
        } else {
            // If tech did not fill any Task field, delete page
            console.println("deleting page " + k);
            this.deletePages(k);
            
            // Reduce page count by 1 every time a page is deleted
            pageCount -= 1;
        }
    }

    t.end();
    
    
    //--------------
    // Reorder pages
    //--------------

    console.println("Start of Reordering pages");

    //Start the progress bar
    var t = app.thermometer;
    pageCount = this.numPages;
    t.duration = pageCount;
    t.begin();
    var p = 0;

    // Filter fieldArray to only UnitIDs and store in filteredUnitIDs
    var filteredUnitIDs = fieldArray.filter( (u) => u.type == "UnitID" );
    
    // create a new array sorted by the UnitID found in the field on the page
    var sortedUnitIDs = filteredUnitIDs.sort( function(a,b) { return a.val.localeCompare(b.val) } );
    
    // Loop through the array and move pages
    console.println("pageCount = " + pageCount);
    
    for (i = 0; i < pageCount - 1; i++) {
        //update thermometer stage
        t.text = "Processing page " + p + " of " + this.numPages;
        p += 1;
        
        var exist = getField(sortedUnitIDs[i].name);
        if (exist == null) {
            continue;
        } else {
            console.println("moving " + getField(sortedUnitIDs[i].name).page + " to " + (i+1));
            this.movePage(getField(sortedUnitIDs[i].name).page, i);
            console.println("Moved the page");
        }
    }

    t.end();

    console.println("End of Reordering pages");
    app.alert("Units have been alphabetized and blank pages removed.", 3, 0, "Cleaning Complete");
}

// Function takes a field name string and returns 1 if at least one task field has been filled out. Returns 0 otherwise. Use as boolean.
function taskingDone(tasks) {
    var count = 0;
    
    //Loop through array
    for (var i = 0; i < tasks.length - 1; i++) {
        if (tasks[i].val != "" && tasks[i].val != "Off") {
            //If Task field is filled in or checked, increase count to 1
            count += 1;
            
            console.println("Tech filled in " + tasks[i].val + " into field " + tasks[i].name);
            
            //If at least one field is checked or filled in, stop counting
            break;
        }
    }

    // Use output as boolean
    return count;
}
