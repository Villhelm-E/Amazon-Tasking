//array for unit types
types = ["AC Split System", "Air Handling/Ducted Split System", "Air-Cooled Chiller", "Exhaust Fan", "Gas Unit/Infrared Heater", "Air Curtain", "RTU", "Make-Up Air Unit", "Energy Recovery Ventillator/Unit"];

//array to translate Amazon's naming scheme for unit type to tech-usable unit type
translated = ["Split System/MS/VRV/VRF/FC", "Split Ducted/Gas/HP/AH", "Air-Cooled Chiller", "Exhaust Fan", "Gas/Infrared Heater", "Air Curtain", "RTU", "Make-Up Air", "Energy Recovery Ventillator/Unit"]

//array to map unit types to the type field in templates
template_names = ["AC_Split_System", "Air_Curtain", "Chiller_Fall", "Chiller_Spring", "Ducted_Split_System", "ERV_ERU", "Exhaust_Fan", "Filters", "Gas_Unit", "MAU_Fall", "MAU_Spring", "RTU_Fall", "RTU_Spring"]

//array for seasonal PMs
pms = ["Major/Spring", "Minor/Fall", "Filter Change/Inspection"];

//Create table array for template lookup
table = [];

// make 2d array where first array is PM
table[0] = [];
table[1] = [];
table[2] = [];
table[3] = [];

// Major/Spring
table[0][0] = "AC_Split_System";
table[0][1] = "Ducted_Split_System";
table[0][2] = "Chiller_Spring";
table[0][3] = "Exhaust_Fan";
table[0][5] = "Air_Curtain";
table[0][6] = "RTU_Spring";
table[0][7] = "MAU_Spring";
table[0][8] = "ERV_ERU";
table[0][9] = "DHU_Spring";

// Minor/Fall
table[1][0] = "AC_Split_System";
table[1][1] = "Ducted_Split_System";
table[1][2] = "Chiller_Fall";
table[1][3] = "Exhaust_Fan";
table[1][4] = "Gas_Unit";
table[1][5] = "Air_Curtain";
table[1][6] = "RTU_Fall";
table[1][7] = "MAU_Fall";
table[1][8] = "ERV_ERU";
table[1][9] = "DHU_Fall"

// Filters
table[2][0] = "Filters";
table[2][1] = "Filters";
table[2][2] = "Filters";
table[2][3] = "Filters";
table[2][4] = "Filters";
table[2][5] = "Filters";
table[2][6] = "Filters";
table[2][7] = "Filters";
table[2][8] = "Filters";
table[2][9] = "Filters";