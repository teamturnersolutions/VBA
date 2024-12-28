# Master Staffing Report


High-level:

1st flow:
*Scheduled
Check folder for file status
	If exisit's archive
	If doesn't exisit, end flow

2nd flow:
*Automated
Download email attachment

3rd flow:
*Automated
* when new file is detected
	Execute office script

4th flow:
*Scheduled (Monday morning)
Send Email of office script output (Pivot table)



Mid-Level:

1. Flow: Archive Old File (Scheduled Weekly)
Trigger: Scheduled (e.g., Saturday evening).
Logic:
Check if a file exists in the target folder (e.g., /MasterStaffing/WeeklyReport.xlsx).
If Exists:
Move the file to the Archive folder (/MasterStaffing/Archive).
Rename the file with a timestamp for uniqueness.
If Doesn't Exist: End the flow.
2. Flow: Download Email Attachment (Automated)
Trigger: When an email with an attachment arrives in the inbox.
Logic:
Check if the attachment contains specific criteria (e.g., file name includes "Master Staffing").
Download the attachment to the target folder (e.g., /MasterStaffing/).
3. Flow: Execute Office Script (Automated)
Trigger: When a new file is added to the folder (e.g., /MasterStaffing/).
Logic:
Detect the new file event.
Execute an Office Script (e.g., run the script that processes the data and generates a pivot table).
Store the processed output in a specific folder (e.g., /MasterStaffing/Processed/).
4. Flow: Send Processed File (Scheduled Monday Morning)
Trigger: Scheduled (e.g., every Monday morning).
Logic:
Locate the processed file (e.g., /MasterStaffing/Processed/WeeklyReport_Pivot.xlsx).
Attach the file to an email.
Send the email to the recipients with a summary or instructions.












1. Notifications and Logging (Best Practices)
Notifications:
Purpose: Keep stakeholders informed of success, failure, or important milestones in the process.
Implementation:
Use Microsoft Teams or email connectors to send status updates.
Example triggers for notifications:
File successfully archived.
New file downloaded and processed.
Flow failure (with error details).
Logging:
Purpose: Track process execution for auditing or troubleshooting.
Implementation:
Use a dedicated Excel or SharePoint list for log entries.
Example log details:
Timestamp of each step.
File name and location of archived/downloaded/processed files.
Success or failure status.
2. UI Layer for Manual Intervention
Why Add a UI Layer?
Simplifies control for non-technical users.
Adds flexibility to rerun, pause, or skip certain steps.
How to Build It:
Use Power Apps as a simple front-end interface.
Features:
Buttons to trigger individual flows (e.g., "Archive Files," "Run Office Script").
Status display (e.g., last run time, current progress).
Error handling (e.g., allow users to retry failed steps).
Access Control: Limit access to specific users while allowing admins to monitor activity.
3. Final Touch: A Gift Presentation
To make this truly feel like a thoughtful gift, package it nicely:

Documentation:
Create a quick guide or user manual for the flows and UI.
Include screenshots and a simple explanation of how the automation benefits them.
Demo Session:
Offer a quick walkthrough to showcase how to use it.
Highlight the notifications and logs for transparency.
Backup and Support Plan:
Mention that you've backed up the flows and are available for tweaks.










Checklist:

Handoff Checklist
1. Duplication in End User’s Environment
Flow Recreation: Rebuild each flow under their credentials or organizational account.
Connections:
Re-establish any linked services (e.g., OneDrive, Office 365, Outlook) using their credentials.
Verify permissions for any external systems or services.
Environment-Specific Testing:
Perform a test run in their setup to confirm compatibility and functionality.
2. Documentation
Written:
Include detailed steps for maintaining or modifying the flows.
Document the logic of each flow with clear descriptions of triggers, conditions, and actions.
Provide troubleshooting tips for common issues.
Recorded:
Walk through the system, highlighting key processes and the purpose of each flow.
Record how to handle notifications, logging, or failures.
Keep it concise and accessible.
3. Ownership and Access
Ownership Transfer:
Make the end user the owner of the flows and any related resources (e.g., SharePoint folders, OneDrive locations).
Remove yourself from owner roles after confirming everything is functional.
Backup Plan:
Save a copy of your work in a safe location (e.g., a personal backup account or local archive).
Inform the end user about the backup and its location.
4. Training and Support
Training:
Schedule a short training session to ensure they’re comfortable managing the system.
Answer any questions and explain how they can adjust flows if needed.
Support:
Offer a follow-up period (e.g., 1-2 weeks) for additional assistance.
Ensure they have a point of contact if issues arise after you step away.
5. Cleanup
Your Accounts:
Ensure no lingering dependencies exist on your credentials.
Logs and Data:
Transfer any important logs or historical data to their environment.
Redundancies:
Delete any duplicates or temporary files in your environment once the handoff is complete.


MOD: cut all unnecessary data. Only data needed for presentation. Reason, Optimized for speed, efficency and reliability

Master Headers
("NAME", "DEPT", "SHIFT", "STATUS")

Department Mapping Table-

"DC Laydown", "DC PTC"
"Tiers", "DC PTC"
"Breakdown", "DC PTC"
"Laydown", "DC PTC"
"PTC Laydown", "DC PTC"
"DC Cross Dock", "DC Cross Dock & Shipping"
"Crossdock", "DC Cross Dock & Shipping"
"DC Ship/Security", "DC SCC"
"Security", "DC SCC"
"DC-E-COMM", "DC E-COMM"
"Maintenance", "DC Maintenance"
"Pallet Land", "DC Pallet Department"
"Palletland", "DC Pallet Department"
"Dc Qc/Lp", "DC QC/LP"
"QC", "DC QC/LP"
"Storage", "DC Storage"
"Truck Audit", "DC Yard Jockey"
"Receiving", "DC Receiving"
    
    
----------------------------------------------------------------------------------------------------------------
ACTIONS TO TAKE ON TABLE QUERIES
----------------------------------------------------------------------------------------------------------------

BJS TMS - 
DELETE COLUMNS:
ID, JOB TITLE, SUPERVISOR/MANAGER, LOCATION, LOA, COLUMN1 
ADD COLUMNS:
STATUS



LGS -
DELETE COLUMNS:
ID, POSITION, START DATE, REASON
RENAME COLUMN:
LAST, FIRST NAME = NAME
DEPARTMENT = DEPT
ADD COLUMNS:
STATUS
