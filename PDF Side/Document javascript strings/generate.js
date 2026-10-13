function generateTasking() {
    load_units();

    ////////////
    // Choose PM
    ////////////
    // Get PM name from field
    chosen_pm = getField("pm").value;
    
    // Get the PM index from the array
    pm_index = pms.indexOf(chosen_pm);

    //////////////
    // Choose Site
    //////////////
    // Filter sites based on Site dropdown
    var chosenSite = units.filter(function (e) {
        return e.Site == getField("customer").value;
    });

    ////////////////////////////////
    // Special PM cases
    ////////////////////////////////

    //If PM chosen is Filter Change/Inspection, filter units to only those with filters
    if (chosen_pm == "Filter Change/Inspection") {
        chosenSite = chosenSite.filter(function(e) {
            return e.Filter1Qty != "";
        });
    }

    ////////////////
    // Prep for loop
    ////////////////

    // count the number of units
    var count = chosenSite.length;
    // console.println("number of units: " + count);

    // Start the progress bar
    var t = app.thermometer;
    t.duration = count;
    t.begin();
    
    ////////////////////////////////////////////////
    // loop through all units in site chosen by user
    ////////////////////////////////////////////////
    for (var i = 0; i < count; i++) {
        //////////////////////
        // Do all the lookups
        //////////////////////
        // Get unit's type
        var unit_type = chosenSite[i].UnitType;
        // Get unit's type index
        var type_index = translated.indexOf(unit_type);

        // update thermometer stage
        t.value = i;
        t.text = "Generating PM " + (i + 1) + " of " + count + " (" + Math.round((i+1)/count*100*100)/100 + "%)";

        // if the unit is a Gas Unit/Infrared Heater during a Spring PM, skip it to avoid error
        if (pm_index == 0 && type_index == 4) { continue; }
        
        // lookup the unit type and pm combination 
        var template_name = table[pm_index][type_index];

        // get page number
        var page = this.numPages;
        // Get unit ID
        var unit_ID = chosenSite[i].UnitID;
        // Get Category
        var category = chosenSite[i].Category;
        // Get unit manufacturer
        var unit_manuf = chosenSite[i].Manufacturer;
        // Get unit model
        var unit_model = chosenSite[i].Model;
        // Get unit serial
        var unit_serial = chosenSite[i].Serial;
        
        // Get drive type
        var drive_type = chosenSite[i].Belt1;
        // The dropdown uses "Belt" as a catch-all for any belt
        if (drive_type == "DD" || drive_type == "" || drive_type == "No access") {
            // do nothing
            // To-do: rewrite this in the future so it's not awkward
        } else {
            drive_type = "Belt";
        }

        // Get Filter 1
        var filter_1 = chosenSite[i].Filter1;
        var filter_1_qty = chosenSite[i].Filter1Qty;
        // Get Filter 2
        var filter_2 = chosenSite[i].Filter2;
        var filter_2_qty = chosenSite[i].Filter2Qty;
        ///////////////////////////////////////////////
        // copy the template to the end of the document
        ///////////////////////////////////////////////

        var template_index = template_names.indexOf(template_name);
        this.templates[template_index].spawn(this.numPages, true, false);
        
        ///////////////
        // fill in form
        ///////////////
        
        // Unit ID
        var unit_pattern = "P"+page+"."+template_name+".UnitID";
        var unit_field = getField(unit_pattern);
        unit_field.value = unit_ID;

        // Unit Type
        var type_pattern = "P"+page+"."+template_name+".UnitType";
        getField(type_pattern).value = unit_type;

        // Category
        var cat_pattern = "P"+page+"."+template_name+".SubType";
        var cat_field = getField(cat_pattern);
        if (category != "") {
            cat_field.value = category;
        }

        // Manufacturer
        var manuf_pattern = "P"+page+"."+template_name+".manufacturer";
        var manuf_field = getField(manuf_pattern);
        manuf_field.value = unit_manuf;
        
        // Model
        var model_pattern = "P"+page+"."+template_name+".model";
        var model_field = getField(model_pattern);
        model_field.value = unit_model;
        
        // Serial
        var serial_pattern = "P"+page+"."+template_name+".serial";
        var serial_field = getField(serial_pattern);
        serial_field.value = unit_serial;
        
        // Drive Type
        if (drive_type != "") {
            var drive = "P"+page+"."+template_name+".DriveType";
            var drive_field = getField(drive);
            drive_field.value = drive_type;
        }

        // Filters
        if (template_name == "Filters") {
            var filter1_pattern = "P"+page+"."+template_name+".filter1";
            var filter1_field = getField(filter1_pattern);
            filter1_field.value = filter_1;
            var filter1Qty_pattern = "P"+page+"."+template_name+".filter1qty";
            var filter1Qty_field = getField(filter1Qty_pattern);
            filter1Qty_field.value = filter_1_qty;
            var filter2_pattern = "P"+page+"."+template_name+".filter2";
            var filter2_field = getField(filter2_pattern);
            filter2_field.value = filter_2;
            var filter2Qty_pattern = "P"+page+"."+template_name+".filter2qty";
            var filter2Qty_field = getField(filter2Qty_pattern);
            filter2Qty_field.value = filter_2_qty;
        }
    }
    
    ///////////
    // Clean up
    ///////////
    // Hide the Generate button and update equipment list button
    getField("Submit").hidden = true;
    getField("reload").hidden = true;

    // Make Site and PM fields read-only
    getField("customer").readonly = true;
    getField("pm").readonly = true;

    // Show the Clean Tasking and Finalize buttons
    getField("Clean").hidden = false;
    getField("Finalize").hidden = false;

    // update the generated variable
    generated = true;

    t.end();
}