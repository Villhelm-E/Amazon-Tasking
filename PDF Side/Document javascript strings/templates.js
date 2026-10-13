function hideTemplates() {
    this.getTemplate("AC_Split_System").hidden = true;
    this.getTemplate("Ducted_Split_System").hidden = true;
    this.getTemplate("Chiller_Spring").hidden = true;
    this.getTemplate("Chiller_Fall").hidden = true;
    this.getTemplate("Exhaust_Fan").hidden = true;
    this.getTemplate("Gas_Unit").hidden = true;
    this.getTemplate("Air_Curtain").hidden = true;
    this.getTemplate("Filters").hidden = true;
    this.getTemplate("RTU_Spring").hidden = true;
    this.getTemplate("RTU_Fall").hidden = true;
    this.getTemplate("MAU_Spring").hidden = true;
    this.getTemplate("MAU_Fall").hidden = true;
    this.getTemplate("ERV_ERU").hidden = true;
}

function showTemplates() {
    this.getTemplate("AC_Split_System").hidden = false;
    this.getTemplate("Ducted_Split_System").hidden = false;
    this.getTemplate("Chiller_Spring").hidden = false;
    this.getTemplate("Chiller_Fall").hidden = false;
    this.getTemplate("Exhaust_Fan").hidden = false;
    this.getTemplate("Gas_Unit").hidden = false;
    this.getTemplate("Air_Curtain").hidden = false;
    this.getTemplate("Filters").hidden = false;
    this.getTemplate("RTU_Spring").hidden = false;
    this.getTemplate("RTU_Fall").hidden = false;
    this.getTemplate("MAU_Spring").hidden = false;
    this.getTemplate("MAU_Fall").hidden = false;
    this.getTemplate("ERV_ERU").hidden = false;
}