// Form Studio — Application State

// =============================================================
// APP STATE
// =============================================================
const AppState = {
  currentView: "home",           // sidebar nav key
  currentUser: null,             // { displayName, email, id }
  isAdmin: false,
  hasFormRequestAccess: false,

  // Form builder wizard state
  builderMode: "create",        // "create" | "edit"
  builderItemId: null,          // SharePoint item ID (when editing)
  builderStep: 0,               // current wizard step
  builderForm: {
    title: "",
    listName: "",
    sections: [],               // [{ id, title, stepIndex, fields: [] }]
    layout: "single",           // "single" | "multistep"
    access: "StaffStudents",
    specificPeople: [],         // [{id, displayName, email}]
    formManagers: [],           // [{id, displayName, email}] — colleagues who can manage submissions
    submissionType: "Submit",
    submitNotifyEmails: "",   // comma-separated — notified when any user submits
    notifySubmitter:    true, // send submitter a read-only HTML confirmation email
    conditions: [],             // [{showFieldId, whenFieldId, equalsValue}]
    dependentDropdowns: [],     // [{childFieldId, parentFieldId, mapping: {parentVal: [childVals]}}]
    governance: {
      existingProcess:    "",   // "yes" | "no" | "" — is this based on an existing process?
      existingProcessDetail: "", // free text — describe the existing process
      retention:          "",   // "under1" | "1to3" | "3to7" | "indefinite" | ""
      sensitiveData:      "",   // "none" | "personal" | "commercial" | "both" | ""
      privacyAssessment:  "",   // "yes" | "no" | "na" | ""
      externalAccess:     "",   // "none" | "recipients" | "submitters" | ""
      continuityPlan:     "",   // free text — workaround if form stops working
      expectedVolume:     "",   // "low" | "medium" | "high" | ""
      dataOwner:          null, // { id, displayName, email } | null — SP Person column
    },
  },

  // Lists data
  allRequests: [],
  liveForms: [],
};

function resetBuilderForm() {
  AppState.builderMode = "create";
  AppState.builderItemId = null;
  AppState.builderStep = 0;
  AppState.builderForm = {
    title: "",
    listName: "",
    sections: [],
    layout: "single",
    access: "StaffStudents",
    specificPeople: [],
    formManagers: [],
    submissionType:     "Submit",
    submitNotifyEmails: "",   // comma-separated — notified when any user submits
    notifySubmitter:    true, // send submitter a read-only HTML confirmation email
    conditions: [],
    dependentDropdowns: [],
    governance: {
      existingProcess:       "",
      existingProcessDetail: "",
      retention:             "",
      sensitiveData:         "",
      privacyAssessment:     "",
      externalAccess:        "",
      continuityPlan:        "",
      expectedVolume:        "",
      dataOwner:             null, // { id, displayName, email } | null
    },
  };
}
