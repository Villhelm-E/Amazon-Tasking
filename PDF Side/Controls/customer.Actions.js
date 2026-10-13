var site = getField("customer").value; 

function SiteLookup() {

    //filter sites based on dropdown
    var chosenSite = units.filter(function (el) {
        return el.Site == site;
    });
    
    //Get site address
    var site_address = chosenSite[0].Address;

    //Get site type
    var site_type = chosenSite[0].SiteType;
    
    //rename address field
    getField("address").value = site_address;

    //rename type field
    getField("site_type").value = site_type;
}

SiteLookup();