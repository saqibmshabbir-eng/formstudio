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
  STAFF_GROUP:   "00000000-0000-0000-0000-000000000000", // TODO: replace with Entra Object ID for FLM-SPO-Staff-Intranet-Visitors
  STUDENT_GROUP: "00000000-0000-0000-0000-000000000000", // TODO: replace with Entra Object ID for FLM-SPO-Student-Intranet-Visitors

  // Column value used to identify admins in FormAdmins list
  ADMIN_TITLE_VALUE: "Admin",

  // ── SharePoint internal column names on the Forms list ───────────────────
  COL_STATUS:   "Status",         // Single line text
  COL_LISTNAME: "ListName",       // Single line text — name of the provisioned data list
  COL_FORM_DEF: "FormDefinition", // Multiline plain text — the JSON blob

  // Microsoft Graph API base URL
  GRAPH_BASE: "https://graph.microsoft.com/v1.0",

  // MSAL scopes required
  SCOPES: [
    "Sites.ReadWrite.All",
    "Sites.Manage.All",
    "Files.ReadWrite.All",
    "User.Read",
    "People.Read",
  ],

  // Supported field types for the form builder
  FIELD_TYPES: [
    { value: "Text",        label: "Single Line Text" },
    { value: "Note",        label: "Multiline Text" },
    { value: "RichText",    label: "Rich Text" },
    { value: "Number",      label: "Number" },
    { value: "Currency",    label: "Currency" },
    { value: "DateTime",    label: "Date / Time" },
    { value: "Boolean",     label: "Yes / No" },
    { value: "Choice",      label: "Choice (Dropdown)" },
    { value: "MultiChoice", label: "Multiple Choice (Checkboxes)" },
    { value: "User",        label: "Person or Group" },
    { value: "Lookup",      label: "Lookup" },
    { value: "URL",         label: "Hyperlink" },
    { value: "Location",    label: "Location" },
    { value: "Thumbnail",   label: "Image" },
    { value: "Calculated",  label: "Calculated" },
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

  // Full lifecycle status flow — single list, single flow
  STATUS_FLOW: {
    "Draft":              [],
    "Submitted":          ["Approve for Review", "Reject"],
    "Approved for Preview": ["Approve", "Reject"],
    "Preview":            ["Approve", "Reject"],
    "Approved":           [],
    "Live":               [],
    "Rejected":           [],
  },

  STATUS_ACTION_MAP: {
    "Approve for Review": "Approved for Preview",
    "Approve":              "Approved",
    "Reject":               "Rejected",
  },
};
