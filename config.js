// Form Studio — CONFIG
// Edit these values before deploying

const CONFIG = {
  // Azure AD / Entra ID settings
  TENANT_ID: "aebecd6a-31d4-4b01-95ce-8274afe853d9",
  CLIENT_ID: "fb5c58c3-9515-4058-a072-d79e134f05e2",

  // SharePoint root site for form storage
  SITE_URL: "https://uniofleicester.sharepoint.com/sites/rooms2",

  // Single SharePoint list covering the full form lifecycle
  // ACTION REQUIRED: Create a list named "Forms" with these columns:
  //   Title          (single line text — built-in)
  //   Status         (single line text)
  //   ListName       (single line text)
  //   FormDefinition (multiline text, plain — NOT rich text)
  FORMS_LIST:       "Forms",       // Replaces both FormRequest and LiveForms
  FORM_ADMINS_LIST: "FormAdmins",  // Admin access control list (unchanged)

  // M365/Entra group Object IDs for permissions on provisioned data lists.
  // Find these in Entra admin centre → Groups → <group> → Overview → Object ID.
  // The c:0t.c|tenant| claim format resolves an Entra group by Object ID via
  // SharePoint's ensureuser endpoint — no directory read permissions required.
  STAFF_GROUP:   "d4b4a155-7c91-4b84-95b9-8b951a695c00", // Entra Object ID for FLM-SPO-Staff-Intranet-Visitors
  STUDENT_GROUP: "01e83af2-cf5c-4788-a433-b930b85d2fb3", // Entra Object ID for FLM-SPO-Student-Intranet-Visitors

  // Column value used to identify admins in FormAdmins list
  ADMIN_TITLE_VALUE: "Admin",

  // ── SharePoint internal column names on the Forms list ───────────────────
  COL_STATUS:        "Status",         // Single line text
  COL_LISTNAME:      "ListName",       // Single line text — name of the provisioned data list
  COL_FORM_DEF:      "FormDefinition", // Multiline plain text — the JSON blob
  COL_COMMENTS:      "AdminComments",  // Multiline plain text — admin approval/rejection notes
  COL_IS_DELETED:       "IsDeleted",       // Yes/No — soft delete flag on data lists
  COL_ASSIGNED_TO:      "AssignedTo",      // Person — soft check-out: who currently "owns" a submission for editing
  COL_ASSIGNED_TO_EMAIL:"AssignedToEmail", // Text — email of the assigned user; written by the app alongside AssignedTo
  COL_SUBMISSION_STATUS:"SubmissionStatus",// Text — submission processing status: Submitted, Processed & Approved, Processed & Declined
  COL_RETRO:         "Retro",          // Yes/No — marks an externally-owned SP list form
  COL_LIST_LOCATION: "ListLocation",   // Single line text — SP new item URL for retro forms
  COL_VIEW_URL:      "ViewUrl",        // Single line text — URL to view submissions for retro forms

  // Governance columns — promoted to SP list columns for admin filtering and reporting
  // ACTION REQUIRED: Add these columns to the Forms list in SharePoint:
  //   GovRetention    (single line text)
  //   GovSensitive    (single line text)
  //   GovPrivacy      (single line text)
  //   GovExternal     (single line text)
  //   GovVolume       (single line text)
  //   GovDataOwner    (Person or Group — single selection, allow only internal users)
  COL_GOV_RETENTION:    "GovRetention",
  COL_GOV_SENSITIVE:    "GovSensitive",
  COL_GOV_PRIVACY:      "GovPrivacy",
  COL_GOV_EXTERNAL:     "GovExternal",
  COL_GOV_VOLUME:       "GovVolume",
  COL_GOV_DATA_OWNER:   "GovDataOwner", // Person column — written as GovDataOwnerLookupId
  SUBMITTER_ROLE: "Form Submitter", // Custom SP role: AddListItems + ViewListItems (own items only)

  // Microsoft Graph API base URL
  GRAPH_BASE: "https://graph.microsoft.com/v1.0",

  // Set to true to enable verbose Graph API error logging in the browser console
  DEBUG_LOGGING: true,
  // Base URL of this app — used to construct deep links in notification emails.
  // Must match the URL users access the app from. No trailing slash.
  APP_URL: "https://saqibmshabbir-eng.github.io/formstudio/",

  SCOPES: [
    "Sites.ReadWrite.All",
    "Sites.Manage.All",
    "Files.ReadWrite.All",
    "User.Read",
    "People.Read",
    "Mail.Send",   // Required for section Complete button — sends notification email as the logged-in user
  ],

  // Supported field types for the form builder
  // Note: Lookup (requires target list), Thumbnail/Image (Graph API limitation),
  // and Calculated (requires formula) are intentionally excluded — they cannot
  // be created programmatically via the Graph API column endpoint.
  FIELD_TYPES: [
    { value: "Text",        label: "Single Line Text" },
    { value: "Note",        label: "Multiline Text" },
    { value: "RichText",    label: "Rich Text" },
    { value: "Number",      label: "Number" },
    { value: "Currency",    label: "Currency" },
    { value: "DateTime",    label: "Date" },
    { value: "Boolean",     label: "Yes / No" },
    { value: "Choice",      label: "Choice (Dropdown)" },
    { value: "MultiChoice", label: "Multiple Choice (Checkboxes)" },
    { value: "User",        label: "Person or Group" },
    { value: "URL",         label: "Hyperlink" },
    { value: "FileUpload",  label: "File Upload" },
    { value: "InfoText",    label: "Info / Notice (display only)" },
  ],

  SUBMISSION_TYPES: [
    { value: "Submit",     label: "Submit Only",   desc: "User submits once; cannot edit after" },
    { value: "SubmitEdit", label: "Submit & Edit", desc: "User can return and edit their submission" },
  ],

  ACCESS_OPTIONS: [
    { value: "StaffStudents", label: "Staff + Students" },
    { value: "StaffOnly",     label: "Staff Only" },
    { value: "Specific",      label: "Specific People Only" },
  ],

  // Field labels that clash with system-managed columns. Checked at builder
  // validation time after sanitiseColumnName, case-insensitive — so "Assigned To",
  // "assigned-to", "AssignedTo!" all collapse to "assignedto" and are blocked.
  // AssignedTo and IsDeleted are added by the provisioner; Title/ID/Status are
  // SharePoint reserved or already-used names on the data list.
  RESERVED_FIELD_NAMES: ["AssignedTo", "AssignedToEmail", "IsDeleted", "Title", "ID", "Status"],

  // Admin-only status actions — defines which action buttons appear in the admin
  // review table for a form in a given status. Does not cover user-driven transitions
  // (Created → Submitted → Created) which are hardcoded directly in builder.js.
  STATUS_FLOW: {
    "Created":            [],
    "Submitted":          ["Approve for Preview", "Reject"],
    "Preview":            ["Approve", "Reject"],
    "Approved":           [],
    "Live":               [],
    "Closed":             [],
    "Rejected":           [],
  },

  STATUS_ACTION_MAP: {
    "Approve for Preview": "Preview",
    "Approve":             "Approved",
    "Reject":              "Rejected",
  },
};
