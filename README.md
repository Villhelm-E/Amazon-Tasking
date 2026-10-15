# Amazon-Tasking
Workflow for maintaining equipment lists in Excel and generating Tasking PDF documents.

## Preventative Maintenance Tasking
Our contracts with clients include periodic Preventative Maintenance procedures (PMs). Each type of unit, such as exhaust fans or air curtains, require a different checklist of maintenance tasks (what we call Tasking). We generated Tasking PDFs for our technicians to perform with our ERP, however in early 2025 Amazon became a new client and with them came new requirements and hundreds of sites.

## The Problem
Our ERP was generating summarized Tasking, where every unit with identical checklist was grouped onto a single page. Amazon required each unit to have it's own Tasking page. The Tasking our ERP generated was also not a fillable form; we had been operating on the honor system by just believing the tech performed every task. Amazon required more concrete proof and so the form had to be fillable.

It wasn't reasonable to expect our office staff to manually add checkboxes to every Tasking generated. We were dealing with 150 sites (later 300), ranging from 5 to 200 units each, with 9 to 31 checkboxes for each unit, generated 4 times a year.

I took on the challenge of automating the process.

## The Old Workflow
### Creating the Equipment List
1. Technicians filling out an equipment list for a site would do so in a PDF.
2. The technician would send the finished equipment to the office, which used it as the current equipment list.
### Generating the Tasking
3. The office would enter the equipment list into our ERP.
4. Once entered in the ERP, it would generate a non-fillable PDF and email the Tasking to a tech.
5. The tech would perform all the tasks but because it wasn't fillable, there was no paper trail.

## The New Workflow
### Creating the Equipment List
1. The office sends the technician/technicians an Equipment List Template. This is spreadsheet that allows technicians to easily copy/paste, auto-increment, and sort units. It's also an Excel Template file, so it couldn't be accidentally overwritten. I also added conditional formatting so the technician can double-check their work.
2. The tech sends the equipment list to the office. The office copies the contents into their Master Equipment List. This is a different spreadsheet that contains all sites the office is charged with.
### Generating the Tasking
3. I added a custom ribbon button to the Master Equipment List that exports the equipment list to a JSON file. The office needs to press this button whenever they make changes/updates to their master equipment list.
4. The user opens the Tasking Generator. This is a PDF that imports the JSON file exported in the previous step. The user chooses the site from a drop-down, chooses the type of PM, enters the SO, and then hits the Generate Tasking button. This button generates all the checklists for each unit at the site selected as its own page.
5. The user saves the PDF and sends to the technician to fill out.
6. The technician fills out the fillable PDF form and sends it back to the office.
7. If multiple technicians filled out the Tasking for a single site, there is a process the office needs to follow to combine them into one Tasking document.
8. Once the Tasking document is cleaned up, they can upload to Amazon's portal.

## The Implementation
I worked with a technician to test out the Equipment List Template. I learned the quirks of the iPad version of Excel and what was unreasonable to ask of a technician. This informed how to design the template. In theory it could also be filled out by multiple techs simultaneously through OneDrive, but we ran into sync issues with poor cell service so we don't use that feature.

I did video training sessions with the Service Manager, Service Supervisor, one technician from each office and every Service Manager and dispatcher at every office.

During the transition period, I had to convert their PDF equipment lists to spreadsheets, and in many cases fix their equipment lists and generate their Tasking for them. Amazon liked the design of my checklists so much they said they want to make it a required format for all Amazon sites.

I created a manual with screenshots of all the steps the office needs to follow.

## Planned Features
I have three more features to add to the Master Equipment List. Currently, updating to new versions of the Master Equipment List spreadsheet is unacceptably involved. The process goes:
1. I send out the new version of the spreadsheet to everyone
2. They have to copy the existing master equipment list to the new spreadsheet without breaking the conditional formatting
3. Save the new master equipment list
4. Delete the old spreadsheet
5. Rename their equipment list
The first feature I plan to add, which I am currently coding, will allow the user to push a button in the ribbon and choose a master equipment list to import and automatically delete the old file.

The second feature I am working on will let the user click a button to import a tech's template equipment list. This will reduce the amount of manual work users currently have to do as well as reduce user error.

The third feature will be another custom ribbon button that will perform automated cleanup of the equipment list such as add lead zeros and alphabetize units.

## End of Life
The company is planning on replacing our ERP with a modern one in the near future. This new ERP will make this whole project obsolete, so I won't be adding new features for very long.
