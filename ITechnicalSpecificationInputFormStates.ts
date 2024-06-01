import { TermActionsDisplayStyle } from "@pnp/spfx-controls-react";
import { DateFormat } from "@pnp/spfx-controls-react/lib/controls/dynamicForm/dynamicField/IDynamicFieldProps";

export interface ITechnicalSpecificationInputFormState {
    currentItem: any;
    globalID: number;
    type: string;
    currentUser: any;
    fileID: number;
    is_idco: string;
    allIDCOOptions: any,
    initiativeDocuments: any;
    errorMessage: string;
    title: string;
    description: string;
    documentNumber: string;
    initiativeID: number;
    adminUser: any;
    comments: string;
    revision: string;
    created: any;
    delegated:any;
    modified: any;
    createdBy: any;
    requestedBy:any;
    modifiedBy: any;
    status: string;
    docCreated: any;
    folderId: number;
    isSMEEvaluationRequired: any;
    Requestforexpiration: boolean;
    isWorkflowInProgress:any;
    isOtherFunctionApprovalRequired:any;

    dropdownRevisions: any;
    dropdownDocumentTypeOptions: any;
    dropdownIDCOOptions: any;
    dropdownDepartmentOptions:any;

    approvalHistorySMEEvaluation: any;
    approvalHistorySMEEvaluationColumns: any;
    approvalHistorySME: any;
    approvalHistorySMEColumns: any;
    approvalHistoryOtherFunctions: any;
    approvalHistoryOtherFunctionsColumns: any;
    approvalHistoryCX: any;
    approvalHistoryCXColumns: any;

    previousRequesterId:any,
    selectedRequester:any;
    selectedDocumentType: number;
    selectedDepartment:number;
    selectedIDCO: number;
    selectedIDCOObj: any;
    selectedSME: any,
    defaultSME: any,
    selectedAttachmentFile: string,
    initiativeDocumentArray: any;
    discussions:any;
    discussionComment:string;

    selectedOFApproverGME:any,
    selectedOFApproverGP:any,
    selectedOFApproverGQ:any,
    selectedOFApprover4:any,
    selectedOFApprover5:any,
    selectedOFApprover6:any,
    insertOFApprovers:boolean,

    isAuthorLoggedIn: any;
    isAdminLoggedIn: any;
    isSMELoggedIn: any;
    isOtherFunctionApproverLoggedIn: any;
    isCXApproverLoggedIn: any;
    currentApprovalHistoryItemOtherFunctions: any;
    currentApprovalHistoryItemCX: any;

    approvalItem: any;
    approvalDone: any;
    smeApprovalItem: any,
    otherFunctionApprovalItem: any,
    cxApprovalItem: any,
    smeApprovalDone: boolean,
    otherFunctionApprovalDone: boolean,
    cxApprovalDone: boolean,
    reviewComments: string,

    isShowSMEEvaluationTab: boolean,
    isShowSMETab: boolean,
    isShowOtherFunctionsTab: boolean,
    isShowCXTab: boolean,

    isShowLoader: boolean,
    isLoading: boolean,
    isNewStage: boolean,
    isNewRevStage: boolean,
    isRequestedStage: boolean;
    isDraftStage: boolean;
    isNotAuthorised: boolean;
    isUnderSMEEvaluation: boolean;
    isSMEEvaluationDone: boolean;
    isSMEEvaluationRejected: boolean;
    isInApprovalSME: boolean;
    isSMEApproved: boolean;
    isSMERejected: boolean,
    isInApprovalOtherFunction: boolean;
    isOtherFunctionsApproved: boolean;
    isOtherFunctionsRejected: boolean;
    isInApprovalCX: boolean;
    isCXApproved: boolean;
    isCXRejected: boolean;
    isSentForClosureApproved: boolean;
    isSentForClosureRejected: boolean;
    isSentForClosureExpired: boolean;
    isClosed:boolean;
    isClosedApproved: boolean;
    isClosedRejected: boolean;
    isClosedExpired: boolean;
    isClosedCancelled:boolean;
    isShowUploadOption: boolean;
    isDisableEditing: boolean;
    isDisableSMESelection:boolean;
    isDisableIDCOSelection:boolean;

    requesterPersona: any,
    editorPersona: any,
    docAuthorPersona: any,
    docEditorPersona: any,
    documentCardActivityPeople: any,
    docPreviewProps: any

}
export interface IUserDetails {
    ID?: number;
    loginName: string;
    email: string;
}
export interface ISpfxPnpPeoplepickerState {
    SuccessMessage: string;
    UserDetails: IUserDetails[];
    selectedusers: string[];
}