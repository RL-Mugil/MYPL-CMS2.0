# Metayage Portal User Manual

## Document Control

- Product: Metayage IP & Patent Services Portal
- Audience: Internal staff, attorneys, galvanizers, admins, super admin, client-side users
- Deployment model: GitHub Pages frontend, Clerk authentication, Cloudflare Worker proxy, Google Apps Script backend, Google Sheets data store, Google Drive document storage
- Manual scope: Operating manual for the current live portal behavior
- Manual purpose: Explain how to use the portal safely, consistently, and efficiently in daily operations

## How To Use This Manual

This manual is an operating guide. It is not written as a developer specification. Read it in three passes:

1. Read the portal overview, login, navigation, roles, and common rules.
2. Read the modules relevant to your role.
3. Read the workflows and examples that match your daily work.

Super Admin and Admin should read the entire manual. Galvanizers should read the queue, workflow board, notifications, approvals, documents, and organization sections carefully. Staff should read cases, documents, inbox, threads, daily audit, task usage, and expense claims. Attorneys should read attorney workspace, inbox, threads, workflow board, and case handling sections. Client-facing roles should focus on client, document, and organization-linked visibility sections.

## Portal Purpose

The Metayage Portal is the operating system for patent-service execution. It centralizes:

- users and roles
- clients and organizations
- cases and workflow stages
- document access and uploads
- approvals and document requests
- direct chats and structured threads
- notifications
- daily audit and wrap-up reporting
- expense claims
- dashboard monitoring

The portal is designed to replace fragmented tracking across spreadsheets, email, phone calls, and informal chat by giving each matter a clear operational path.

## Portal Architecture In Plain Language

The browser UI is hosted on GitHub Pages. Users authenticate with Clerk. The frontend sends requests to a Cloudflare Worker. The Worker forwards those requests to Google Apps Script. Google Apps Script runs the business logic and reads and writes data in Google Sheets and Google Drive.

What this means operationally:

- the system is easy to deploy and cost-effective
- some pages may be slower than a conventional SaaS product if they request too much live data at once
- lookup data, dashboard data, and direct inbox data are cached where safe
- document and bill files are stored in Drive, not inside the frontend

## Core Concepts

### Client Code

Each client has a code. This is a key search value across the system.

- India clients normally end with `M`
- Abroad clients normally end with `Y`

Examples:

- `A61M`
- `870Y`

Always search by client code when possible. It is the fastest and cleanest way to locate matters.

### Case ID

Each case ID is generated from the client code and the next running sequence.

Example:

- client code `A61M`
- next sequence `002`
- case ID `A61M002`

### Organization vs Individual Client

There are two client models:

- `Individual`
- `Organization`

If the client type is `Individual`:

- organization fields do not apply
- client admin user mapping does not apply

If the client type is `Organization`:

- the client can be linked to an organization record
- organization-linked external users can see organization-related client data and documents, subject to role limits

### Internal Users vs External Users

Internal users:

- Super Admin
- Admin
- Galvanizer
- Staff
- Attorney

External users:

- Client Admin
- Client Employee
- Individual Client

External users must not see internal finance unless explicitly permitted by role and business rule.

## Roles In The Portal

### Super Admin

Highest internal authority.

Typical powers:

- full system visibility
- user, client, organization, circle, and case oversight
- expense claim review
- approval oversight
- dashboard oversight
- internal governance

### Admin

Day-to-day operational controller.

Typical powers:

- manage users
- manage clients
- manage organizations
- run bulk updates
- monitor dashboard
- review daily audit entries
- manage work distribution

### Galvanizer

Operational workflow controller between staff and attorney stages.

Typical powers:

- review and route case readiness
- manage galvanizer queue
- use command center view
- push matters toward attorney review
- manage organization profile edits
- manage circles

### Staff

Execution-oriented internal user.

Typical powers:

- create and update matters
- manage client information
- upload documents
- communicate in inbox and threads
- submit daily priorities and wrap-up
- submit expense claims

### Attorney

Legal review role.

Typical powers:

- use attorney workspace
- review assigned cases
- communicate via inbox and threads
- track assigned legal tasks

### Client Admin

Highest role on client company side.

Typical powers:

- view organization-linked records allowed to client-side access
- manage external client-side visibility within company context
- coordinate with internal team

### Client Employee

Company-side user under a client organization.

Typical powers:

- view the client/company-specific records made visible to that organization
- view documents and case data tied to their organization
- finance is restricted unless business rules explicitly allow it

### Individual Client

Single-user client access model, not organization-based.

Typical powers:

- view their own client-side case and document information
- limited to their own matters

## Login And Session Behavior

### Login

Users log in through Clerk. After successful login, the portal exchanges the identity with the backend and maps the user against `USER_ROLES`.

If login works but portal access fails, common causes are:

- email not mapped in `USER_ROLES`
- inactive status
- wrong role mapping

### Session Expiry

If the session expires:

- the portal shows a session-expired or backend error
- the user must log in again

### Sign Out

Sign out clears browser-side cache related to portal lookups and session-linked data.

## Navigation

The left navigation is the main operating menu.

### Sidebar Controls

Users can:

- collapse the full sidebar
- expand or collapse `Portfolio`
- expand or collapse `Operations`

This is useful when working with large tables or the inbox.

### Main Sections

- Dashboard
- Cases
- Documents
- Finance
- Inbox
- Threads
- Notifications
- Daily Audit
- Expense Claims
- Galvanizer Queue
- Management

Depending on role, some sections may be hidden.

## Theme Switching

The portal supports both dark and light themes.

### How Theme Switching Works

- dark theme is the default
- a theme toggle is available in the top bar
- the user can switch between dark and light at any time
- the selection is stored in browser local storage

### What Is Theme-Aware

The following are designed to remain readable in both themes:

- body background
- cards
- borders
- forms
- tables
- text and muted text
- badges
- inbox and chat surfaces
- dashboard charts and chart labels

## Dashboard

The dashboard gives the current operational snapshot.

### Dashboard Cards

The summary cards show:

- Granted
- Pending
- Deadlines
- Unpaid
- Unread Alerts
- Open Threads
- My Clients

### Dashboard Filters

Dashboard filters are inside the filter button. The filter drawer can include:

- Client Code
- Case ID
- From Date
- To Date

After applying filters:

- the drawer closes
- summary and detail sections refresh

### Dashboard Charts And Detail Area

The dashboard loads in two stages:

1. summary cards
2. charts and detailed widgets

This improves perceived speed.

### What The Charts Mean

- Granted Patents: granted matters by country
- Pending Patents: active pending matters by country
- Pending by Status: active pending cases grouped by current status
- Upcoming Renewals: nearest deadlines
- Pending Payments: unpaid invoices or pending invoice items
- Action Required: recent active matters needing attention

## Cases

Cases are the central matter records.

### Case List

The list can be filtered by:

- Client Code
- Case ID
- Search text
- Status
- Country
- Workflow Stage

The system uses `Load More` to reduce heavy first-load delay.

### New Case Form

The new case form uses searchable inputs where structured data already exists.

Searchable fields include:

- Client Code
- Assigned Staff
- Galvanizer
- Attorney
- Organization ID

### How Organization Works In Cases

If organization is selected:

- organization-linked visibility can apply to client-side organization users

If organization is not selected:

- the case still belongs to the client and can inherit organization through the client record

### Case Assignment Notifications

When a case is assigned or reassigned to:

- Staff
- Galvanizer
- Attorney

the assignee receives a notification.

### Bulk Update

Admins can use bulk update to modify multiple cases together for fields like:

- staff
- galvanizer
- attorney
- workflow stage
- status
- priority
- next deadline

## Workflow Board

The workflow board gives a stage-based view of cases.

Typical columns:

- Drafting
- Ready for Attorney
- Under Attorney Review
- Filed
- Under Examination
- Granted
- Closed

### Why It Exists

The cases table is good for records. The workflow board is good for flow.

Use it to answer:

- where are matters currently stuck
- which matters are ready to move
- which stage is overloaded

## Galvanizer Queue And Command Center

The galvanizer module is the operational control zone for moving matters from internal preparation toward legal review and onward.

### Queue Purpose

The queue focuses on matters in stages such as:

- Drafting
- Ready for Attorney
- Under Attorney Review

### Command Center Summary

The command center gives summary counts such as:

- Incoming
- Ready for Attorney
- Under Review
- Pending Approvals

### Galvanizer Flow Example

Example 1:

1. Staff creates or updates a case.
2. Staff completes drafting inputs and document preparation.
3. Workflow stage is set to `Ready for Attorney`.
4. Galvanizer reviews whether:
   - client code is correct
   - matter details are complete
   - documents are available
   - right attorney is assigned
5. If ready, the matter moves to `Under Attorney Review`.

Example 2:

1. A matter is incomplete.
2. Galvanizer sees missing document or unclear note.
3. Galvanizer sends direct message or thread instruction.
4. Galvanizer may create a task or document request.
5. Matter stays in an earlier stage until corrected.

### Good Galvanizer Practice

- do not move a matter forward without required documents
- use notifications and messages instead of out-of-band verbal memory
- use approvals for sensitive transitions when needed

## Attorney Workspace

This view is for attorney-oriented work.

It typically shows:

- pending review matters
- active review matters
- attorney-linked tasks

Use it to avoid scanning the whole case list.

## Inbox

The inbox is for person-to-person direct chat.

### Current Model

The inbox is designed as a two-pane direct messaging view:

- left: user/chat list
- right: active chat panel

### Inbox Purpose

Use Inbox for:

- direct internal discussions
- quick coordination
- one-to-one follow-up
- short operational messages

### How It Works

- open Inbox
- choose a user from the left panel
- the right panel opens the conversation
- type in the composer and send

### Recommended Use

Use Inbox for:

- quick coordination
- direct clarifications
- person-specific updates

Do not use Inbox when the matter needs structured, case-linked, or client-linked audit visibility. Use Threads instead.

## Threads

Threads are structured conversations.

### Thread Purpose

Use Threads when the conversation should remain tied to a work context:

- general internal thread
- case-linked discussion
- client-linked discussion
- internal-only issue

### Difference Between Inbox And Threads

Inbox:

- direct
- person-to-person
- quick chat style

Threads:

- structured
- title-based
- context-linked
- better for auditable collaboration

### Thread Soft Delete

When a thread is deleted from the portal:

- it disappears from the active frontend
- it is soft-deleted in the backend
- it remains in sheets for audit/history

## Notifications

Notifications show system events relevant to the logged-in user.

### Common Notification Triggers

- case assignment
- new direct message
- new thread message
- approval request
- approval decision
- expense claim changes
- document request updates
- daily priority or wrap-up events where applicable

### Notification Actions

Users can:

- mark a notification read
- clear a single notification
- clear all notifications

Clearing is soft delete. Cleared notifications vanish from the portal but remain available in backend history.

## Tasks

Tasks convert work into assignable action items.

Task fields can include:

- title
- description
- assignee
- priority
- due date
- related client or case
- notes

Use tasks for work that must be tracked, not just discussed.

## Activity Timeline

The activity timeline is the operational event feed.

Typical events:

- case updates
- task events
- approval events
- document request events
- thread or message events

Use the timeline when you want to understand what happened and when.

## Approvals

Approvals are used when one user requests a formal decision from another.

### Approval Lifecycle

1. Request approval
2. Approver receives notification
3. Approver opens approval screen
4. Approver selects approve or reject
5. Requester receives decision notification

### Typical Uses

- case-stage approval
- finance-sensitive approval
- document-related approval
- internal decision control

## Documents

The Documents section is the main document viewing and upload area.

### Current Capabilities

- list documents by category
- filter by client code using searchable dropdown
- upload documents directly from the portal
- store uploaded files in the selected client’s Drive folder/category

### Common Categories

- Applications
- Office Actions
- Responses
- Certificates
- Invoices
- Communication

### Upload Flow

1. Open Documents
2. Choose client code
3. Choose category
4. Choose file
5. Upload

This helps place documents into the correct client folder instead of uploading them manually into Drive.

## Document Workflow

Document Workflow is different from the general Documents page.

### Difference

Documents:

- actual file view and upload

Document Workflow:

- request and track required documents
- assign document-related work
- follow request status

### Document Request Use Cases

- signed POA required
- missing company authorization
- inventor document pending
- client submission awaited

## Finance

Finance contains invoice-related visibility and controls. Access depends on role and finance permissions.

Client-side organization users should not see finance unless explicitly allowed by system role rules.

## Expense Claims

Expense Claims are internal-only.

### New Claim

Users can submit:

- claim date
- category
- amount
- bill image upload or bill link
- description

### Bill Upload

The portal supports direct bill upload for:

- `.jpg`
- `.jpeg`
- `.png`

The uploaded bill is stored in Drive and the `Open` action opens it.

### History Filters

History supports:

- This Month
- All Time
- From Date
- To Date

These filters are shown in one compact row.

## Daily Audit

Daily Audit is the internal accountability module.

### Daily Priorities

Users enter:

- Priority 1
- Priority 2
- Priority 3
- Notes

### Day-End Wrap-Up

Users enter:

- High points
- Low points
- Help needed

### Admin Review Use

Admins and above can review:

- what the user planned
- what the user reported later
- where help was needed

### Filters

Daily Audit uses compact filters:

- Date
- Name

Name is selected as a user-facing value instead of email for better practice.

## Management

Management is for administration and setup control.

### Users

User creation uses structured inputs wherever possible.

Important current rules:

- Department is a dropdown:
  - Management
  - External
- Organization ID is searchable
- if organization is selected for a user:
  - department becomes `External`
  - role becomes client-side role logic
  - finance visibility is disabled unless business logic changes it

### Additional Roles

Additional roles should be assigned using structured controls, not free typing.

### Clients

Client forms behave differently by type.

If type is `Individual`:

- organization fields should not apply
- client admin user fields should not apply

If type is `Organization`:

- organization linkage applies
- client admin mapping can apply
- assigned staff can be chosen from searchable dropdown

### Organizations

Organizations are maintained from the Management > Organizations screen.

Current behavior:

- new organizations are created with an in-portal form
- Galvanizer and above can edit organization profile
- each organization can have linked clients
- the organization screen is used to attach existing clients to the organization, not to create new users directly

### How To Add To Organization

Flow:

1. Management
2. Organization tab
3. Open a specific organization
4. Click `Add Client`
5. Search existing clients in the popup
6. attach the correct client

The popup search is interactive. It filters while typing and does not require exact-name-only search.

### Organization Visibility

If a user is linked under an organization on the external side, organization-linked data and documents become visible according to role rules. Finance remains restricted unless explicitly allowed.

## Circles

Circles are internal grouping constructs.

### What Circles Do

- group internal users
- support internal coordination and navigation grouping

### Delete Behavior

Circle deletion is soft delete.

When a circle is deleted:

- it vanishes from the portal
- it remains in backend data for traceability

## Smart Search

Smart Search helps find records across:

- cases
- clients
- tasks
- messages/threads
- users where access allows

Use it for fast navigation when you know part of the value but not the exact location.

## User Guidance By Role

### Super Admin Daily Routine

Recommended sequence:

1. review dashboard
2. review notifications
3. review approvals
4. review expense claims
5. inspect management exceptions

### Admin Daily Routine

Recommended sequence:

1. dashboard
2. notifications
3. daily audit
4. management updates
5. case assignments or bulk updates

### Galvanizer Daily Routine

Recommended sequence:

1. dashboard
2. galvanizer queue
3. workflow board
4. approvals
5. direct inbox or threads for clarifications

### Staff Daily Routine

Recommended sequence:

1. notifications
2. inbox
3. tasks
4. cases
5. documents
6. daily priorities
7. wrap-up before end of day

### Attorney Daily Routine

Recommended sequence:

1. notifications
2. inbox
3. attorney workspace
4. case-linked thread review
5. task and deadline follow-up

### Client Admin Daily Routine

Recommended sequence:

1. review organization-linked client records
2. review visible documents
3. check notifications and threads
4. coordinate with internal team on pending needs

## Common Workflows

### Workflow 1: Create A New Client And Case

1. Create client in Management > Clients
2. confirm client code
3. if organization-based, link organization
4. create case using client code
5. assign staff, galvanizer, attorney as needed
6. assigned users receive notifications

### Workflow 2: Move Matter To Attorney Review

1. Staff updates case details
2. documents are uploaded
3. galvanizer reviews readiness
4. workflow stage moves to `Ready for Attorney`
5. attorney assignment is confirmed
6. attorney receives notification

### Workflow 3: Request A Missing Document

1. create document request
2. assign owner or target
3. monitor notification and status
4. upload actual file in Documents once available

### Workflow 4: Internal Clarification

Use Inbox if the message is person-to-person.

Use Threads if:

- it belongs to a case
- it belongs to a client
- it needs structured title/context

### Workflow 5: Expense Claim Submission

1. open Expense Claims
2. fill claim details
3. upload image or add bill link
4. submit
5. approver reviews later

## Troubleshooting

### Login Works But Portal Access Fails

Check:

- user exists in `USER_ROLES`
- email matches exactly
- status is active
- role is mapped correctly

### Dashboard Seems Slow

The system uses summary-first loading and caching, but first load can still be slower because of Apps Script and Sheets. Repeat loads should be faster.

### Inbox Or Threads Do Not Show Expected Messages

Check:

- you are using Inbox for direct messages
- you are using Threads for structured discussions
- the users involved have permission to see the conversation

### Organization User Visibility Seems Wrong

Check:

- whether the user is an internal role or external role
- whether the client is linked to the organization
- whether the role is client admin or client employee
- finance expectations must remain restricted unless explicitly allowed

### Notification Did Not Disappear

Use:

- mark read, or
- clear, or
- clear all

Clearing is soft delete and should remove it from the portal view.

## Data Entry Rules

Wherever structured data already exists, use searchable dropdowns or controlled inputs rather than free text.

Examples of structured-entry fields:

- client code
- case ID
- organization ID
- user email
- client admin user ID
- assigned staff email
- attorney
- galvanizer

Free text should be limited to areas like:

- address
- notes
- descriptions
- message text
- document titles where needed
- external links when not handled by upload

## Good Operating Practices

- search by client code before creating a new client
- search by case ID before editing a case
- use Inbox for direct chat, Threads for auditable structured discussions
- upload documents from the portal to the correct client/category whenever possible
- use approvals for decision points, not informal assumptions
- keep daily priorities and wrap-up consistent
- clear notifications after acting on them

## Final Notes

The portal is designed to support disciplined operational behavior. It works best when users:

- use searchable structured fields instead of free typing
- keep assignments accurate
- attach records to the correct client and organization
- use the right communication module for the right purpose
- maintain daily reporting discipline

This is not just a display system. It is a workflow and accountability system. The more consistently the team uses the portal as intended, the more useful the dashboard, queue, notifications, approvals, documents, and history become.
