function onlyUnique(value, index, array) {
  return array.indexOf(value) === index;
}

function uniqueSites(array) {
    var c = [];
    for (var i = 0; i < array.length; i++) {
        c[i] = array[i].Site;
    }
    return c.filter(onlyUnique);
}

// function to load a combo/list box with an array of values
function LoadOptions(oField, aValues) {
    var bResult = false;
    // load array of values into field object
    if (oField.type != "combobox" && oField.type != "listbox") {
        app.alert("Load Options function requires a combobox or listbox", 0, 0, "Combobox Not Found");
        
        bResult = false;
    } else {
        oField.setItems(aValues); // set values
        bResult = true;
    }

    return bResult;
}

function reload_dropdown() {
    var uniques = uniqueSites(units);
    LoadOptions(this.getField('customer'), uniques);
}