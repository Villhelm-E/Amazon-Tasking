# Amazon-Tasking
Workflow for conducting equipment surveys, maintaining equipment lists, and generating PDF checklists for our contracts with Amazon.

## Preventative Maintenance *Tasking*
Our contracts with clients include periodic *Preventative Maintenance* procedures (PMs). Each type of unit, such as exhaust fans or air curtains, require a different checklist of maintenance tasks (*Tasking*). The *Tasking* PDFs our ERP creates are not fillable; we operated on the honor system. Amazon became a customer in 2025 and required checkboxes on the *Tasking*. It was not reasonable to expect our office staff to manually add checkboxes to every *Tasking* generated. We were dealing with 150 sites, ranging from 5 to 200 units each, with 9 to 31 checkboxes for each unit, generated 4 times a year. A new workflow was needed with the following requirements:
### Requirements
1. The final document must look presentable for Amazon.
2. The workflow must be usable from beginning to end without IT needing to be involved.
3. Assume the end user has low technical ability.
4. Tasks and equipment lists must be fillable by technicians in the field.
5. The sytem must handle split surveys and split *PMs*.
6. Each unit's *Tasking* must be on its own page.
7. The system must generate spring, fall, annual, and belt *PMs*

Ideally, this would be a web app but I don't know how to build a web app yet. However, I do know some VBA and Javascript so I started a feasibility study and then determined it was possible.

## The Old Workflow
### Creating the Equipment List
1. Technicians filled out equipment lists on a PDF on their iPads.
2. The technician would email the finished equipment list to the office, which would reference for the most current.
### Generating the *Tasking*
3. The office would enter the equipment list into our ERP.
4. Once entered in the ERP, it would generate a non-fillable PDF and email the *Tasking* to a tech.
5. The tech would perform all the tasks but because it wasn't fillable, there was no paper trail.

## The New Workflow
### Creating the Equipment List
1. The office sends the technician/technicians an Equipment List Template. This is a spreadsheet that allows technicians to easily copy/paste, auto-increment, and sort units. It's also an Excel Template file, so it can't be accidentally overwritten. I also added conditional formatting so the technician can double-check their work.
2. The tech sends the equipment list to the office. The office copies the contents into their Master Equipment List. This is a different spreadsheet that contains all sites the office is responsible for.
### Generating the Tasking
3. The user presses the ribbon button in the Master Equipment List that exports the equipment list to a JSON file. The office needs to press this button whenever they make changes/updates to their master equipment list.
4. The user opens the Tasking Generator. This is a PDF that imports the JSON file exported in the previous step. The user chooses the site and PM type from a drop-down, enters the SO, and then hits the Generate Tasking button. This button generates a page for each unit in the equipment list with the appropriate *Tasking* for the unit type and PM type.
5. The user saves the PDF and sends to the technician to fill out.
6. The technician fills out the form and sends it back to the office.
7. If multiple technicians filled out the *Tasking* for a single site, there is a process the office needs to follow to consolidate the *Tasking* documents.
8. Once the *Tasking* document is cleaned up, they can upload to Amazon's portal.

## The Implementation
I worked with technicians and Operations Manager to understand the *PM* scopes and cetegorize the unit types, which informed the *Tasking* templates. For a few months I had to convert their PDF equipment lists to spreadsheets and generate all the *Tasking* while I figured out how to make it user-friendly. I worked with a technician to test out the Equipment List Template on iPad in the field, which informed the design of the template. In theory the spreadsheet could also be filled out by multiple techs simultaneously through OneDrive, but we ran into sync issues with poor cell service and abandoned that feature.

Once the first user-friendly version was ready, I hosted video training sessions with the Service Supervisor; Operations Manager; all Service Managers and Dispatchers; and at least one technician from each office. I took feedback to improve the system and created a manual for the whole process. Amazon liked the design of my *Tasking* so much they want to make it a required format for all Amazon sites.

## Planned Features
Automated cleanup of the equipment list such as adding lead zeros and alphabetizing units

## End of Life
The company is planning on moving to a modern ERP in the near future, which will make this whole workflow obsolete. Thus I won't be making many more updates.