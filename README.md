# PowerShell-Scripting

<h1>Production PowerShell Scripting</h1>

<b> This repository is a list of all current production PowerShell scripts I have created and implemented for DDB WorldWide over the past two years</b>

<p> You will find the descrptions of each script in this README file. Each script is ran inside of its own Scheduled Task on a single server. The files themselvs are available in this repository for purusal. Please contact me at danjfrank08@gmail.com if you have any questions!</p>

<b>Notes:</b>

<ol>
	<li>For Office 365 scripts, “DDBNA” defines DDB North America Office 365 tenants.</li>
	<li>“DDBWW” defines DDB Worldwide Office 365 tenants. Included tenants are noted in each report email.</li>
	<li>For Local AD, “DDBNA” defines DDB North America domains</li>
	<li>All run times listed are PST.</li>
	<li>An automated script runs weekly that archives reports older than 30 days and deletes reports older than 90 days.</li>
</ol>

<h2>Office 365 Scripts</h2>

<h3>DDB North America (DDBNA)</h3>

<pre>
Name: DDBNA O365 Complete Report
Description: Details complete mailbox and user information for DDB North American tenants.

The following data is collected:
Display Name
First Name
Last Name
User Principal Name
Email Address
License Type
Office
Department
User Disabled
MFA Status
MFA Excluded
User Forward SMTP
User Internal Forward
User SMTP Deliver and Forward
Retention Policy
Retention Hold Enabled
Litigation Hold Enabled
Litigation Hold Owner
Tenant
Tenant Office Name
Office 365 Primary Domain               
Cloud Only Account
Mailbox Type
Run Time: Monday 12:00 AM 

Name: DDBNA O365 Mailbox Auditing
Description: Enables auditing on recently created mailboxes in North American Office 356 tenants.
Run Time: Monthly:  2:00 AM on the 1st

Name: DDBNA O365 Retention Mailbox Size Report
Description: Report emailed monthly on Retention mailbox sizes for each office 365 tenant.
Run Time: Monthly:  1:00 AM on the 1st

Name: DDBNA O365 MFA Users To Enable
Description: A helper script that creates a list of users whom are to be enabled within the next week. Emails are sent to these users, and a list of users within 30 days from creation date are emailed to site admins.
Run Time: Tuesday 6:00 AM

Name: DDBNA O365 User Enabled Tomorrow Notice
Description: The list created by the "Helper_DDBNA_O365_MFA_Enablement" script is run through and emails are sent to remind users they will be enabled the following day.
Run Time: Monday 6:00 AM

<h3>DDB Worldwide (DDBWW)</h3>

Name: DDBWW O365 Complete Report
Description: Details complete mailbox and user information for DDB Worldwide tenants.

The following data is collected:
Display Name
First Name
Last Name
User Principal Name
Email Address
License Type
Office
Department
User Disabled
MFA Status
MFA Excluded
User Forward SMTP
User Internal Forward
User SMTP Deliver and Forward
Retention Policy
Retention Hold Enabled
Litigation Hold Enabled
Litigation Hold Owner
Tenant
Tenant Office Name
Office 365 Primary Domain               
Cloud Only Account
Mailbox Type
Run Time: Sunday 11:00 PM 

Name: DDBWW O365 Domains Report
Description: A report listing out the connected DDB World Wide offices' domains.
Run Time: Monday 7:00 AM

Name: DDBWW O365 Retention Policy Report
Description: Ran against all available DDB WW offices. The DDB WW Retention Policy Report tells us what retention policies are available and which tags are associated with the policies.
Run Time: Monthly: 3:30 AM on the 1st

<h2>Local AD Scripts</h2>
 
Name: DDBNA Local AD Password Expiration Notices
Description: Script to run against all DDB NA ADs to check for password expiration. Notify user their AD password will need to be updated. Starts 15 days out from expiration, emails user daily at 8:00 AM PST
Run Time: Daily 8:00 AM 

Name: DDBNA Local AD Password Expiration Summary
Description: Reports on which passwords are expired and when they expired for DDB North America.
Run Time: Monday 9:00 AM

Name: DDBNA Local AD Freelancer Report
Description: Report on all non-disabled freelancers across North America.
Run Time: Tuesday 11:00 AM

Name: DDBNA Local AD Workstation Cleanup
Description: Runs against all SBOX domains (excluding Interbrand) in DDB North America. Moves any workstations in the Computers OU to the Workstation OU. Removes Workstations from AD that have not been logged into for more than 120 days.
Run Time: Monday 2:00 AM

Name: DDBNA Local AD Disabled Users Cleanup
Description: Removes disabled users older than 30 days in DDB North American ADs.
Run Time: Daily 7:00 AM
</pre>
