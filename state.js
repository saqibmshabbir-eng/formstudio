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
    conditions: [],             // [{showFieldId, whenFieldId, equalsValue}]
    dependentDropdowns: [],     // [{childFieldId, parentFieldId, mapping: {parentVal: [childVals]}}]
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
    submissionType: "Submit",
    conditions: [],
    dependentDropdowns: [],
  };
}
