# csi-thingymajig
CSI thingymajig for CSRs.

## Overview
- Purpose
	- The main purpose of this program is to assist those who have trouble organizing their thoughts and help create a flow that keeps the call in an enclosed environment, and in the end, create a CSI note.
- Workflow summary
	1. Agent opens the program alongside other CSR tools.
	2. Agent does their usual call flow/greeting.
	3. Agent inputs the necessary information needed in the fields for the task/concern to be resolved.
	4. Premade scripts are given and the process of resolving the concern is shown.
	5. Agent follows steps given from program.
	6. Once task is resolved, a CSI is given for the Agent to copy and paste in their CSI tool.
	7. The program resets the fields for the next caller.

---
## VBA Form Design
- Controls used
	- **Main Form**
		- Text boxes
			- Caller Name - Agent can input caller name here for premade scripts.
			- Caller Subscriber ID - policy ID/Subscriber ID of caller.
			- Comment Box - Agent can use the extra text box to write multiple concerns or comments from the caller.
		- Combo box
			- Concern - a dropdown list that contains common concerns callers have.
			- Plan - a dropdown list that contains plans of GRIC.
		- Checkboxes
			- First Name - Used for CIC Verification. Required to unlock Load button. 
			- Last Name - Used for CIC Verification. Required to unlock Load button.
			- Birthdate - Used for CIC Verification. Required to unlock Load button.
			- ~~Mailing Address - Used for CIC Verification.~~
			- Physical Address - Used for CIC Verification. Required to unlock Load button.
		- Buttons
			- Load - Redirects the agent how to solve the selected concern and plan. Is greyed out until CIC verification is done.

	- **Adding Dependent Form**
		- Text boxes
		- Combo box
			- Dependent Relationship - Agent can select dependent relationship. For CSI use later.
			- Plan - Agent can select plan to add dependent. For CSI use later.
		- Buttons
			- Load - Shows the necessary steps to add dependent based on the plan.
- Naming conventions
	- Variables are named using lower camelCase.

---
## Script Templates
- Table of concerns + scripts
	- Adding Dependent
	- Appeals & Grievances
	- Cancellation
	- Reactivation
	- Reapplication
	- Removing Dependent
	- Reinstatement
	- Request VOC
	- Termination
	- Withdrawal
- Placeholder logic

---
## Branching Logic
- Flowchart of resolution paths
- Notes per concern type

---
## Future Ideas
- Save call logs
- Add retention offers
- Migrate to Python or Power Apps
