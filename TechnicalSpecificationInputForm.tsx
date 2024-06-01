import * as React from 'react';
import style from './TechnicalSpecificationInputForm.module.scss';
import { ITechnicalSpecificationInputFormProps } from './ITechnicalSpecificationInputFormProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { ITechnicalSpecificationInputFormState } from './ITechnicalSpecificationInputFormStates';
import { spListNames } from './ITechnicalSpecificationInputFormProvider';
import TechnicalSpecificationInputFormProvider, { ITechnicalSpecificationInputFormProvider } from './ITechnicalSpecificationInputFormProvider';

import Moment from 'react-moment';
Moment.globalFormat = 'DD/MM/YYYY hh:mm A';
import { Fabric, isDark, Stylesheet, Toggle } from 'office-ui-fabric-react';
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from '@fluentui/react/lib/Persona';
import { containsInvalidFileFolderChars, sp } from "@pnp/sp/presets/all";
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { TextField, MaskedTextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { Checkbox, DetailsList, DetailsListLayoutMode, DocumentCard, DocumentCardActivity, DocumentCardDetails, DocumentCardPreview, DocumentCardTitle, DocumentCardType, ExpandingCard, ExpandingCardMode, FontWeights, HoverCard, HoverCardType, IColumn, IDocumentCardActivityPerson, IDocumentCardPreviewProps, IExpandingCardProps, ImageFit, IPlainCardProps, Link, mergeStyles, mergeStyleSets, MessageBar, MessageBarType, Pivot, PivotItem, SelectionMode, Spinner, Stack } from '@fluentui/react';
import { DatePicker, IDatePickerStyles, defaultDatePickerStrings, SpinButton, ISpinButtonStyles, Position } from '@fluentui/react';
import { Separator } from '@fluentui/react/lib/Separator';
import { Icon } from '@fluentui/react/lib/Icon';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CommandBarButton, IconButton, DefaultButton, CommandButton } from '@fluentui/react/lib/Button';
import { Label } from '@fluentui/react/lib/Label';
import 'bootstrap/dist/css/bootstrap.min.css';
import { BaseClientSideWebPart, BaseWebPartContext } from '@microsoft/sp-webpart-base';
import * as moment from 'moment';
import styles from '@pnp/spfx-controls-react/lib/controls/iFrameDialog/IFrameDialogContent.module.scss';
require('./TechnicalSpecificationInputForm.module.scss');

const logo: any = require('../../assets/welcome-dark.png');

/*****************************************************************************************/
//Defining styles for form controls
const itemClasses = mergeStyleSets({
  selectors: {
    '&:hover': {
      textDecoration: 'underline',
      cursor: 'pointer',
      color: "red"
    },
  },
  hoverItem: {
    padding: "10px",
    'span': {
      fontSize: "11px"
    },
    fontSize: "12px"
  },
  label: {
    color: '#041e50',
    cursor: 'pointer'
  },
  cards: {
    maxWidth: '500px'
  }
});
const verticalStyle = mergeStyles({
  height: '200px',
});
const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { maxWidth: "100%", border: "none", borderBottom: "1px solid #C4c3cd" } };
const dropdownStyles: Partial<IDropdownStyles> = { title: { border: "none" }, dropdown: { maxWidth: "100%", border: "none", borderBottom: "1px solid #C4c3cd" } };
const datePickerStyles: Partial<IDatePickerStyles> = { root: { maxWidth: "100%", border: "none", borderBottom: "1px solid #C4c3cd" } };
const spinButtonStyles: Partial<ISpinButtonStyles> = { root: { maxWidth: "100%", border: "none", borderBottom: "1px solid #C4c3cd" } };
const loaderStyles = { root: { maxWidth: "100%", border: "none", Position: "relative", top: "40%" } };
const peoplePickerStyles = { text: { border: "none", borderBottom: "1px solid #C4c3cd" }, fieldGroup: { maxWidth: "100%", border: "none", borderBottom: "1px solid #C4c3cd" } };
const OFpeoplePickerStyles = { text: { border: "none", borderBottom: "1px solid #C4c3cd" }, fieldGroup: { maxWidth: "100%", border: "none", borderBottom: "1px solid #C4c3cd" }, root: { fontSize: "18px" } };
const docCardSTyles = {
  name: { color: "#2e2e38" }, activity: { color: "#2e2e38" }, avatars: { height: "20px", width: "20px" }, avatar: { height: "20px", width: "20px" }, root: {
    marginLeft: "-0px", marginBottom: "5px", avatars: { height: "70px", width: "20px", avatar: { height: "20px", width: "20px" } },
  }
}
const docCardTitleSTyles = { root: { fontSize: "14px", fontWeight: "500" } }
const docCardImgSTyles = {

}
const personaStyles = {
  primaryText: { color: "#fff", ':hover': "#fff", selectors: { ':hover': { color: "#fff" } } }, primaryTextHovered: { color: "#fff", ':hover': "#fff" }, secondaryText: { color: "#fff" }, tertiaryText: { color: "#fff", display: "block", fontSize: "12px" }, root: { color: "#fff!important" }, details: { minHeight: "40px" }
}
const buttonStyles = {
  label: { color: "#041e50", fontWeight: "500" }
}
const pivotStyles = { itemContainer: { padding: "20px 10px" } }
/*****************************************************************************************/

export default class TechnicalSpecificationInputFormComponent extends React.Component<ITechnicalSpecificationInputFormProps, ITechnicalSpecificationInputFormState, {}> {
  private importFileUploadRef: React.RefObject<HTMLInputElement>;
  private _Provider: ITechnicalSpecificationInputFormProvider;
  //Initializing the state parameters and binding functions
  constructor(props: ITechnicalSpecificationInputFormProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context as any
    });
    this.state = {
      currentItem: null,
      globalID: null,
      type: null,
      fileID: null,
      currentUser: null,
      title: null,
      documentNumber: null,
      description: null,
      allIDCOOptions: [],
      is_idco: "No",
      errorMessage: null,
      initiativeDocuments: null,
      initiativeID: null,
      adminUser: null,
      comments: "",
      revision: null,
      created: null,
      delegated: null,
      modified: null,
      createdBy: null,
      requestedBy: null,
      modifiedBy: null,
      docCreated: null,
      status: null,
      folderId: null,
      isSMEEvaluationRequired: "Yes",
      Requestforexpiration: false,
      isWorkflowInProgress: "Yes",
      isOtherFunctionApprovalRequired:"Yes",

      selectedOFApproverGME: null,
      selectedOFApproverGP: null,
      selectedOFApproverGQ: null,
      selectedOFApprover4: null,
      selectedOFApprover5: null,
      selectedOFApprover6: null,
      insertOFApprovers: false,

      approvalHistorySMEEvaluation: null,
      approvalHistorySMEEvaluationColumns: null,
      approvalHistorySME: null,
      approvalHistorySMEColumns: null,
      approvalHistoryOtherFunctions: null,
      approvalHistoryOtherFunctionsColumns: null,
      approvalHistoryCX: null,
      approvalHistoryCXColumns: null,

      dropdownRevisions: [],
      dropdownDocumentTypeOptions: [],
      dropdownIDCOOptions: [],
      dropdownDepartmentOptions: [],
      initiativeDocumentArray: [],

      previousRequesterId: null,
      selectedRequester: null,
      selectedDocumentType: null,
      selectedDepartment: null,
      selectedIDCO: null,
      selectedIDCOObj: null,
      selectedSME: [],
      defaultSME: [],
      selectedAttachmentFile: "",
      discussions: [],
      discussionComment: "",

      requesterPersona: null,
      editorPersona: null,
      docAuthorPersona: null,
      docEditorPersona: null,
      documentCardActivityPeople: null,
      docPreviewProps: null,

      isShowSMEEvaluationTab: false,
      isShowSMETab: false,
      isShowOtherFunctionsTab: false,
      isShowCXTab: false,

      isShowLoader: false,
      isLoading: true,
      isNotAuthorised: false,
      isNewStage: false,
      isNewRevStage: false,
      isRequestedStage: false,
      isDraftStage: false,
      isUnderSMEEvaluation: false,
      isSMEEvaluationDone: false,
      isSMEEvaluationRejected: false,
      isInApprovalSME: false,
      isSMEApproved: false,
      isSMERejected: false,
      isInApprovalOtherFunction: false,
      isOtherFunctionsApproved: false,
      isOtherFunctionsRejected: false,
      isInApprovalCX: false,
      isCXApproved: false,
      isCXRejected: false,
      isSentForClosureApproved: false,
      isSentForClosureRejected: false,
      isSentForClosureExpired: false,
      isClosed: false,
      isClosedApproved: false,
      isClosedRejected: false,
      isClosedExpired: false,
      isClosedCancelled: false,
      isShowUploadOption: false,
      isDisableEditing: false,
      isDisableSMESelection: true,
      isDisableIDCOSelection: true,

      isAuthorLoggedIn: false,
      isAdminLoggedIn: false,
      isSMELoggedIn: false,
      isOtherFunctionApproverLoggedIn: false,
      isCXApproverLoggedIn: false,

      approvalItem: null,
      approvalDone: false,
      smeApprovalItem: null,
      otherFunctionApprovalItem: null,
      cxApprovalItem: null,
      smeApprovalDone: false,
      otherFunctionApprovalDone: false,
      cxApprovalDone: false,
      reviewComments: null,

      currentApprovalHistoryItemOtherFunctions: false,
      currentApprovalHistoryItemCX: false,
    };
    this.importFileUploadRef = React.createRef();
    this._onAttachmentFileUpload = this._onAttachmentFileUpload.bind(this);
    this._onsubmit = this._onsubmit.bind(this);
    this._onsubmit = this._onsubmit.bind(this);
  }
  public componentWillMount(): void {
    this._Provider = new TechnicalSpecificationInputFormProvider();
  }
  //Performing component did mount activities
  public async componentDidMount() {
    try {
      //Getting docuemnt id if present
      this.setState({ globalID: !isNaN(parseInt(this._Provider.getParameterByName("DocID"))) ? parseInt(this._Provider.getParameterByName("DocID")) : null });
      //Getting form control inputs
      const [documentTypeOptions, departmentOptions, adminUsers, IDCOOptions, allIDCOs, currentUser] = await Promise.all([
        this._Provider.getSPLookupColumnValues(spListNames.lst_documentTypes, true),
        this._Provider.getSPLookupColumnValues(spListNames.lst_departments, true),
        this._Provider.getcurrentUserGroups('Admins'),
        //this._Provider.getSPListItems("Category eq 'Admin'", "ID", "*,Responsible/Id,Responsible/EMail", "Responsible", spListNames.lst_configurations),
        this._Provider.getSPLookupColumnValues(spListNames.lst_IDCO, true),
        this._Provider.getSPListItems(null, 'Title', true, '*,Approver/Id,Approver/EMail,Approver/Name,Approver/Title', 'Approver', spListNames.lst_IDCO,),
        this._Provider.getcurrentUserDetails()
      ])
      //Setting form control inputs
      this.setState({
        dropdownDocumentTypeOptions: documentTypeOptions,
        dropdownDepartmentOptions: departmentOptions,
        dropdownIDCOOptions: IDCOOptions,
        allIDCOOptions: allIDCOs,
        currentUser: currentUser,
        adminUser: adminUsers
      });
      //If document id is null, checking for type and setting form states. Otherwise get document details
      if (this.state.globalID === null) {
        this.setState({ insertOFApprovers: true });
        this.setState({ type: this._Provider.getParameterByName("type") });
        if (this.state.type === 'R') {
          this.setState({ isNewRevStage: true });
        }
        else {
          this.setState({ isNewStage: true });
        }
        this.setState({ isLoading: false });
      }
      else {
        await this._getInitiativeOtherFunctionApprovers();
        await this._getInitiativeDetails();
        let [discussions] = await Promise.all([this._Provider.getSPListItems("InitiativeIdId eq  " + this.state.globalID, 'Created', false, '*,Author/Id,Author/EMail,Author/Name,Author/Title', 'Author', spListNames.lst_initiativeDiscussions)])
        //Setting form control inputs
        this.setState({
          discussions: discussions
        });
      }
    }
    catch (error) {
      this._handlingException(error);
    }
  }
  //Get initiative details basis document id and assigning form control values.
  public async _getInitiativeDetails() {
    //Resetting form inputs
    this.setState({
      currentItem: null,
      type: null,
      fileID: null,
      title: null,
      documentNumber: null,
      description: null,
      is_idco: "No",
      errorMessage: null,
      initiativeDocuments: null,
      initiativeID: null,
      comments: "",
      revision: null,
      created: null,
      modified: null,
      createdBy: null,
      requestedBy: null,
      modifiedBy: null,
      docCreated: null,
      status: null,
      folderId: null,
      isSMEEvaluationRequired: true,
      Requestforexpiration: false,
      isWorkflowInProgress: true,
      isOtherFunctionApprovalRequired:"Yes",

      approvalHistorySMEEvaluation: null,
      approvalHistorySMEEvaluationColumns: null,
      approvalHistorySME: null,
      approvalHistorySMEColumns: null,
      approvalHistoryOtherFunctions: null,
      approvalHistoryOtherFunctionsColumns: null,
      approvalHistoryCX: null,
      approvalHistoryCXColumns: null,

      initiativeDocumentArray: [],

      selectedDocumentType: null,
      selectedDepartment: null,
      selectedIDCO: null,
      selectedIDCOObj: null,
      selectedSME: [],
      defaultSME: [],
      selectedAttachmentFile: "",
      discussions: [],
      discussionComment: "",

      requesterPersona: null,
      editorPersona: null,
      docAuthorPersona: null,
      docEditorPersona: null,
      documentCardActivityPeople: null,
      docPreviewProps: null,

      isShowSMEEvaluationTab: false,
      isShowSMETab: false,
      isShowOtherFunctionsTab: false,
      isShowCXTab: false,

      isShowLoader: false,
      isLoading: true,
      isNotAuthorised: false,
      isNewStage: false,
      isNewRevStage: false,
      isRequestedStage: false,
      isDraftStage: false,
      isSMEEvaluationDone: false,
      isUnderSMEEvaluation: false,
      isSMEEvaluationRejected: false,
      isInApprovalSME: false,
      isSMEApproved: false,
      isSMERejected: false,
      isInApprovalOtherFunction: false,
      isOtherFunctionsApproved: false,
      isOtherFunctionsRejected: false,
      isInApprovalCX: false,
      isCXApproved: false,
      isCXRejected: false,
      isSentForClosureApproved: false,
      isSentForClosureRejected: false,
      isSentForClosureExpired: false,
      isClosedApproved: false,
      isClosedRejected: false,
      isClosedExpired: false,
      isClosedCancelled: false,
      isClosed: false,
      isShowUploadOption: false,
      isDisableEditing: false,
      isDisableSMESelection: true,
      isDisableIDCOSelection: true,

      isAuthorLoggedIn: false,
      isAdminLoggedIn: false,
      isSMELoggedIn: false,
      isOtherFunctionApproverLoggedIn: false,
      isCXApproverLoggedIn: false,

      approvalItem: null,
      approvalDone: false,
      smeApprovalItem: null,
      otherFunctionApprovalItem: null,
      cxApprovalItem: null,
      smeApprovalDone: false,
      otherFunctionApprovalDone: false,
      cxApprovalDone: false,
      reviewComments: null,

      currentApprovalHistoryItemOtherFunctions: false,
      currentApprovalHistoryItemCX: false,
    });
    const initiativeItem = await this._Provider.getInitiativeDetails(this.state.globalID);
    let item = initiativeItem[0];
    this.setState({ currentItem: item });
    let m = moment();
    //Settimg form input controls
    this.setState({
      documentNumber: item.DocumentNumber,
      revision: item.Revision,
      title: item.Title,
      description: item.Description,
      selectedDocumentType: item.DocumentTypeId !== null ? item.DocumentTypeId : undefined,
      selectedDepartment: item.DepartmentId !== null ? item.DepartmentId : undefined,
      created: moment(item.Created).format('DD MMM YYYY hh:mm A'),//item.Created,
      //modified: moment(item.Modified).format('DD MMM YYYY hh:mm A'),
      requestedBy: item.Requester,
      createdBy: item.CreatedByUser,
      //modifiedBy: item.Editor,
      delegated: moment(item.Delegatedat).format('DD MMM YYYY hh:mm A'),
      modified: moment(item.LastUpdatedAt).format('DD MMM YYYY hh:mm A'),
      modifiedBy: item.LastUpdatedBy,
      selectedIDCO: item.IDCOId !== null ? item.IDCOId : undefined,
      selectedIDCOObj: item.IDCOId !== null ? { Description: item.IDCODescription, ApproverId: item.IDCOApproverId, ApproverEmail: item.IDCOApprover.EMail } : null,
      selectedRequester: item.RequesterId !== null ? { RequesterId: item.RequesterId, RequesterEmail: item.Requester.EMail, RequesterTitle: item.Requester.Title } : null,
      previousRequesterId: item.RequesterId,
      folderId: item.FolderIDId !== null ? item.FolderID : null,
      is_idco: item.IDCO_x0028_Yes_x002f_No_x0029_ ? "Yes" : "No",
      status: item.Status,
      comments: item.RequestorComments,
      isSMEEvaluationRequired: item.SMEEvaluationRequired ? "Yes" : "No",
      isWorkflowInProgress: item.InitiateItem_UpdateWF ? "Yes" : "No",
      isOtherFunctionApprovalRequired: item.OtherFunctionApprovalRequired ? "Yes" : "No",
      Requestforexpiration: item.Requestforexpiration
    });
    //Setting sme object (people array)
    if (item.SMEId !== null) {
      let smeObj = [];
      let defaultObj = [];
      item.SME.map(item => {
        smeObj.push({ id: item.Id, EMail: item.EMail })
        defaultObj.push(item.EMail);
      });
      this.setState({ selectedSME: smeObj, defaultSME: defaultObj });
    }
    //Generating persona control for requester and editor
    let requsterNameText = '', requestedSecondaryText = '', tertiaryText = '';
    if (this.state.currentItem.RequesterId !== this.state.currentItem.AuthorId) {
      //requsterNameText=this.state.requestedBy.Title;
      requestedSecondaryText = "Delegated at " + this.state.delegated;
      tertiaryText = "Created by " + this.state.createdBy + " at " + this.state.created;
    }
    else {
      requestedSecondaryText = "Created at " + this.state.created;
    }
    let requesterPersonaVar: IPersonaSharedProps = {
      imageUrl: this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&accountname=" + this.state.requestedBy.EMail,
      text: this.state.requestedBy.Title,
      secondaryText: requestedSecondaryText,//"Created at " + this.state.created,
      tertiaryText: tertiaryText,
      size: 12
    };
    this.setState({ requesterPersona: requesterPersonaVar });
    let editorPersonaVar: IPersonaSharedProps = {
      imageUrl: this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&accountname=" + this.state.modifiedBy.EMail,
      text: this.state.modifiedBy.Title,
      secondaryText: "Modified at " + this.state.modified,
      size: 12
    };
    this.setState({ editorPersona: editorPersonaVar })
    this._setDocPreview().then(() => { }).catch(() => { });
    this._defineFormState();
  }
  public async _getInitiativeOtherFunctionApprovers() {
    let otherFunctionApprovers = await this._Provider.getSPListItems("InitiativeIDId eq  " + this.state.globalID, 'Function/Title', true, '*,DocumentType/Title,Function/Title,Approver/Title,Approver/EMail', 'DocumentType,Function,Approver', spListNames.lst_initiativeApproversOtherFunctions);//(filter, orderby, sortAsc, select, expand, listName)
    this.setState({
      selectedOFApproverGME: null,
      selectedOFApproverGP: null,
      selectedOFApproverGQ: null,
      selectedOFApprover4: null,
      selectedOFApprover5: null,
      selectedOFApprover6: null
    });
    let counter = 1;
    if (otherFunctionApprovers.length > 0) {
      this.setState({ insertOFApprovers: false });
    }
    else {
      this.setState({ insertOFApprovers: true });
    }
    otherFunctionApprovers.forEach(reviewerObj => {
      switch (reviewerObj.Function.Title) {
        case "Global Manufacturing Engineering": this.setState({ selectedOFApproverGME: { function: reviewerObj.Function.Title, functionId: reviewerObj.FunctionId, docTypeID: reviewerObj.DocumentTypeId, ApproverId: reviewerObj.ApproverId, ApproverEmail: reviewerObj.Approver.EMail, itemId: reviewerObj.ID } });
          break;
        case "Global Purchasing": this.setState({ selectedOFApproverGP: { function: reviewerObj.Function.Title, functionId: reviewerObj.FunctionId, docTypeID: reviewerObj.DocumentTypeId, ApproverId: reviewerObj.ApproverId, ApproverEmail: reviewerObj.Approver.EMail, itemId: reviewerObj.ID } });
          break;
        case "Global Quality": this.setState({ selectedOFApproverGQ: { function: reviewerObj.Function.Title, functionId: reviewerObj.FunctionId, docTypeID: reviewerObj.DocumentTypeId, ApproverId: reviewerObj.ApproverId, ApproverEmail: reviewerObj.Approver.EMail, itemId: reviewerObj.ID } });
          break;
      }
    });
  }
  public async generateInitiativeOtherFunctionApprovers() {
    let otherFunctionApproversVal = await this._Provider.getSPListItems("DocumentTypeId eq  " + this.state.selectedDocumentType, 'Function/Title', true, '*,DocumentType/Title,Function/Title,Approver/Title,Approver/EMail', 'DocumentType,Function,Approver', spListNames.lst_approversOtherFunctions).then((otherFunctionApprovers) => {
      this.setState({
        selectedOFApproverGME: null,
        selectedOFApproverGP: null,
        selectedOFApproverGQ: null,
        selectedOFApprover4: null,
        selectedOFApprover5: null,
        selectedOFApprover6: null
      });
      //this.setState({ insertOFApprovers: true });
      otherFunctionApprovers.forEach(reviewerObj => {
        switch (reviewerObj.Function.Title) {
          case "Global Manufacturing Engineering": this.setState({ selectedOFApproverGME: { function: reviewerObj.Function.Title, functionId: reviewerObj.FunctionId, docTypeID: reviewerObj.DocumentTypeId, ApproverId: reviewerObj.ApproverId, ApproverEmail: reviewerObj.Approver.EMail, itemId: null } });
            break;
          case "Global Purchasing": this.setState({ selectedOFApproverGP: { function: reviewerObj.Function.Title, functionId: reviewerObj.FunctionId, docTypeID: reviewerObj.DocumentTypeId, ApproverId: reviewerObj.ApproverId, ApproverEmail: reviewerObj.Approver.EMail, itemId: null } });
            break;
          case "Global Quality": this.setState({ selectedOFApproverGQ: { function: reviewerObj.Function.Title, functionId: reviewerObj.FunctionId, docTypeID: reviewerObj.DocumentTypeId, ApproverId: reviewerObj.ApproverId, ApproverEmail: reviewerObj.Approver.EMail, itemId: null } });
            break;
        }
      });
      this.setState({ isShowLoader: false });
    })
      .catch((error) => {
        alert("Something went wrong. Please try again later");
        console.log(error);
        window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
      });

    //this.generatePushOFApproversObject("Insert");
  }
  public async generatePushOFApproversObject(type): Promise<any> {
    if (!this.state.insertOFApprovers) {
      //return true;
      type = "Update";
    }
    let OFItem = null, OFApproversArr = [];
    if (this.state.selectedOFApproverGME !== null) {
      OFItem = null;
      OFItem = {
        InitiativeIDId: this.state.globalID,
        FunctionId: this.state.selectedOFApproverGME.functionId,
        DocumentTypeId: this.state.selectedOFApproverGME.docTypeID,
        ApproverId: this.state.selectedOFApproverGME.ApproverId
      }
      OFApproversArr.push({
        item: OFItem,
        result: null,
        error: null,
        id: this.state.selectedOFApproverGME.itemId
      });
    }
    if (this.state.selectedOFApproverGP !== null) {
      OFItem = null;
      OFItem = {
        InitiativeIDId: this.state.globalID,
        FunctionId: this.state.selectedOFApproverGP.functionId,
        DocumentTypeId: this.state.selectedOFApproverGP.docTypeID,
        ApproverId: this.state.selectedOFApproverGP.ApproverId
      }
      OFApproversArr.push({
        item: OFItem,
        result: null,
        error: null,
        id: this.state.selectedOFApproverGP.itemId
      });
    }
    if (this.state.selectedOFApproverGQ !== null) {
      OFItem = null;
      OFItem = {
        InitiativeIDId: this.state.globalID,
        FunctionId: this.state.selectedOFApproverGQ.functionId,
        DocumentTypeId: this.state.selectedOFApproverGQ.docTypeID,
        ApproverId: this.state.selectedOFApproverGQ.ApproverId
      }
      OFApproversArr.push({
        item: OFItem,
        result: null,
        error: null,
        id: this.state.selectedOFApproverGQ.itemId
      });
    }
    let list = sp.web.lists.getByTitle(spListNames.lst_initiativeApproversOtherFunctions);
    let entityTypeFullName = await list.getListItemEntityTypeFullName().then((entityTypeFullNameValue: any) => {
      return entityTypeFullNameValue;
    });
    let results = await this._Provider.processBatch(list, entityTypeFullName, OFApproversArr, 0, type);
    return { "updatedResults": results };
  }
  public _onOFApproverChangeGME = (items: any): void => {
    let OFApprover = this.state.selectedOFApproverGME;
    if (items.length > 0) {
      items.map(item => {
        if (item.loginName !== null && item.loginName !== undefined) {
          //let newObj={ Id: item.id, EMail: item.loginName.split("|membership|")[1] };
          let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: item.id, ApproverEmail: item.loginName.split("|membership|")[1], itemId: OFApprover.itemId }
          this.setState({ selectedOFApproverGME: newObj });
        }
      });
    }
    else {
      let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: null, ApproverEmail: null, itemId: OFApprover.itemId };
      this.setState({ selectedOFApproverGME: newObj });
    }
  }
  public _onOFApproverChangeGP = (items: any): void => {
    let OFApprover = this.state.selectedOFApproverGP;
    if (items.length > 0) {
      items.map(item => {
        if (item.loginName !== null && item.loginName !== undefined) {
          //let newObj={ Id: item.id, EMail: item.loginName.split("|membership|")[1] };
          let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: item.id, ApproverEmail: item.loginName.split("|membership|")[1], itemId: OFApprover.itemId }
          this.setState({ selectedOFApproverGP: newObj });
        }
      });
    }
    else {
      let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: null, ApproverEmail: null, itemId: OFApprover.itemId };
      this.setState({ selectedOFApproverGP: newObj });
    }
  }
  public _onIDCOChange = (items: any): void => {
    if (items.length > 0) {
      items.map(item => {
        if (item.loginName !== null && item.loginName !== undefined) {
          //this.setState({ selectedRequester: { RequesterId: item.id, RequesterEmail: item.loginName.split("|membership|")[1],RequesterTitle:item.text } });
          let selectedIDCO = this.state.selectedIDCOObj;
          this.setState({
            selectedIDCOObj: { Description: selectedIDCO.Description, ApproverId: item.id, ApproverEmail: item.loginName.split("|membership|")[1] }
          });
        }
      });
    }
    else {
      this.setState({ selectedRequester: null });
    }

  }
  public _onOFApproverChangeGQ = (items: any): void => {
    let OFApprover = this.state.selectedOFApproverGQ;
    if (items.length > 0) {
      items.map(item => {
        if (item.loginName !== null && item.loginName !== undefined) {
          //let newObj={ Id: item.id, EMail: item.loginName.split("|membership|")[1] };
          let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: item.id, ApproverEmail: item.loginName.split("|membership|")[1], itemId: OFApprover.itemId }
          this.setState({ selectedOFApproverGQ: newObj });
        }
      });
    }
    else {
      let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: null, ApproverEmail: null, itemId: OFApprover.itemId };
      this.setState({ selectedOFApproverGQ: newObj });
    }
  }
  public _onOFApproverChange4 = (items: any): void => {
    let OFApprover = this.state.selectedOFApprover4;
    if (items.length > 0) {
      items.map(item => {
        if (item.loginName !== null && item.loginName !== undefined) {
          //let newObj={ Id: item.id, EMail: item.loginName.split("|membership|")[1] };
          let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: item.id, ApproverEmail: item.loginName.split("|membership|")[1], itemId: OFApprover.itemId }
          this.setState({ selectedOFApprover4: newObj });
        }
      });
    }
    else {
      let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: null, ApproverEmail: null, itemId: OFApprover.itemId };
      this.setState({ selectedOFApprover4: newObj });
    }
  }
  public _onOFApproverChange5 = (items: any): void => {
    let OFApprover = this.state.selectedOFApprover5;
    if (items.length > 0) {
      items.map(item => {
        if (item.loginName !== null && item.loginName !== undefined) {
          //let newObj={ Id: item.id, EMail: item.loginName.split("|membership|")[1] };
          let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: item.id, ApproverEmail: item.loginName.split("|membership|")[1], itemId: OFApprover.itemId }
          this.setState({ selectedOFApprover5: newObj });
        }
      });
    }
    else {
      let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: null, ApproverEmail: null, itemId: OFApprover.itemId };
      this.setState({ selectedOFApprover5: newObj });
    }
  }
  public async _setDocPreview() {
    //Generate document preview and modified information
    this.setState({ initiativeDocumentArray: [] });
    //Loop through all documents uploaded
    await this._Provider.getDocument(this.state.globalID)
      .then((docArray) => {
        this.setState({ initiativeDocuments: docArray });
        this.state.initiativeDocuments.forEach(initiativeDoc => {
          let docModifiedDate = moment(initiativeDoc.Modified).format('DD MMM YYYY hh:mm A');
          let previewProps: IDocumentCardPreviewProps = {
            previewImages: [
              {
                name: initiativeDoc.FileLeafRef,
                linkProps: {
                  href: initiativeDoc.EncodedAbsUrl + "?csf=1&web=1",
                },
                previewImageSrc: this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/getpreview.ashx?path=" + initiativeDoc.EncodedAbsUrl,
                imageFit: ImageFit.none,
                height: 100,
              },
            ],
          };
          let profilepicImg = this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&accountname=" + initiativeDoc.Editor.EMail;
          let documentCardActivityPeoples = { name: initiativeDoc.Editor.Title, profileImageSrc: profilepicImg }
          let newObj = { docModified: docModifiedDate, documentInfo: initiativeDoc, documentCardActivityPeople: documentCardActivityPeoples, docPreviewProps: previewProps };
          this.setState(prevState => ({
            initiativeDocumentArray: [...prevState.initiativeDocumentArray, newObj]
          }))
        });
      })
      .catch((error) => {

      });
  }
  public _defineFormState() {
    //Defining form states basis initiative status
    let status = this.state.currentItem.Status;
    let generateHistory = false, generateSMEEvaluationHistory = false, generateSMEHistory = false, generateOtherFunctionsHistory = false, generateCXHistory = false, isSMERole = false, isOtherFunctionsRole = false, isCXRole = false;
    //Checks if author logged in
    if (this.state.currentUser.Id === this.state.currentItem.RequesterId) {
      this.setState({ isAuthorLoggedIn: true });
    }
    //console.log("H");
    //Checks if site admin or doc admin logged in
    this.state.adminUser.forEach(grpName => {
      if (grpName.Title === 'Technical Specification Admins' || grpName.Title === 'R&D FPp Technical Specifications Approval Flow Owners') {
        this.setState({ isAdminLoggedIn: true });
      }
    });
    //Disabling form editing if not author or admin. Basis current user role and form status this will be enabled
    if (!this.state.isAuthorLoggedIn && !this.state.isAdminLoggedIn) {
      this.setState({ isDisableEditing: true });
    }
    //Checking for initiative status
    //If requested loading admin form
    //If draft mode load edit form and enable all controls
    //If in approval flow setting form status accordingly which is used to hide/show/disble form controls.
    //Generate approval history basis initiative status and approval level
    switch (status) {
      case "Requested":
        if (this.state.isAdminLoggedIn) {
          this.setState({ isLoading: false, isRequestedStage: true });
        }
        else {
          this.setState({ isLoading: false, isNotAuthorised: true, isRequestedStage: true });
        }
        break;
      case "Draft": this.setState({ isLoading: false, isDraftStage: true, isShowUploadOption: true, isDisableSMESelection: false, isDisableIDCOSelection: false });
        break;
      case "Under SME Evaluation":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        isSMERole = true;
        this.setState({ isLoading: false, isUnderSMEEvaluation: true, isShowUploadOption: true, isDisableIDCOSelection: false });
        break;
      case "SME Evaluation Completed":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        isSMERole = true;
        this.setState({ isLoading: false, isSMEEvaluationDone: true, isShowUploadOption: true, isDisableIDCOSelection: false });
        break;
      case "SME Evaluation Rejected":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        isSMERole = true;
        this.setState({ isLoading: false, isSMEEvaluationRejected: true, isShowUploadOption: true, isDisableIDCOSelection: false });
        break;
      case "Under Approval (SME)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        isSMERole = true;
        this.setState({ isInApprovalSME: true });
        break;
      case "Approved (SME)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        this.setState({ isSMEApproved: true, isShowSMETab: true, isLoading: false, isDisableIDCOSelection: false });
        break;
      case "Rejected (SME)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        this.setState({ isSMERejected: true, isShowSMETab: true, isLoading: false, isShowUploadOption: true, isDisableIDCOSelection: false });
        break;
      case "Under Approval (Other Functions)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        isOtherFunctionsRole = true;
        this.setState({ isInApprovalOtherFunction: true, isShowSMETab: true, });
        break;
      case "Approved (Other Functions)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        this.setState({ isOtherFunctionsApproved: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isLoading: false });
        break;
      case "Rejected (Other Functions)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        this.setState({ isOtherFunctionsRejected: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isLoading: false });
        break;
      case "Under Approval (CX)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        isCXRole = true;
        generateCXHistory = true;
        this.setState({ isInApprovalCX: true, isShowSMETab: true, isShowOtherFunctionsTab: true });
        break;
      case "Approved (CX)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isCXApproved: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false });
        break;
      case "Rejected (CX)":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isCXRejected: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false });
        break;
      case "Sent for Closure / Approved":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isSentForClosureApproved: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false, isDisableEditing: true });
        break;
      case "Sent for Closure / Rejected":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isSentForClosureRejected: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false, isDisableEditing: true });
        break;
      case "Sent for Closure / Expired":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isSentForClosureExpired: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false, isDisableEditing: true });
        break;
      case "Closed / Approved":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isClosedApproved: true, isClosed: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false, isDisableEditing: true });
        break;
      case "Closed / Rejected":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isClosedRejected: true, isClosed: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false, isDisableEditing: true });
        break;
      case "Closed / Expired":
        generateHistory = true;
        generateSMEEvaluationHistory = true;
        generateSMEHistory = true;
        generateOtherFunctionsHistory = true;
        generateCXHistory = true;
        this.setState({ isClosedExpired: true, isClosed: true, isShowSMETab: true, isShowOtherFunctionsTab: true, isShowCXTab: true, isLoading: false, isDisableEditing: true });
        break;
      case "Closed / Cancelled":
        //generateHistory = true;
        //generateSMEEvaluationHistory = true;
        //generateSMEHistory = true;
        //generateOtherFunctionsHistory = true;
        //generateCXHistory = true;
        this.setState({ isClosedCancelled: true, isClosed: true, isLoading: false, isDisableEditing: true });
        break;

    }
    //Check if in approval flow
    if (generateHistory) {
      //Get history items and filter and group the history data basis approval level and hide/show approval tabs
      this._Provider.getSPListItems("InitiativeIDId eq  " + this.state.globalID, 'ID', false, '*,PendingWith/Id,PendingWith/EMail,PendingWith/Name,PendingWith/Title,CXFunction/Title,OtherFunctions/Title', 'PendingWith,CXFunction,OtherFunctions', spListNames.lst_approvalHistory,)
        .then((approvalHistory) => {
          if (generateSMEEvaluationHistory) {
            let smeEvaluationHistory = approvalHistory.filter(his => { return (his.Role === 'SME Evaluation') });
            if (smeEvaluationHistory.length !== 0) {
              this.setState({ isShowSMEEvaluationTab: true });
            }
            this._generateApprovalHistory('SME Evaluation', smeEvaluationHistory, spListNames.lst_approvalHistory);
            if (isSMERole) {
              this._setCurrentUserRole(smeEvaluationHistory, "SME");
            }
          }
          if (generateSMEHistory) {
            let smeApprovalHistory = approvalHistory.filter(his => { return (his.Role === 'SME Approval') });
            if (smeApprovalHistory.length !== 0) {
              this.setState({ isShowSMETab: true });
            }
            this._generateApprovalHistory('SME Approval', smeApprovalHistory, spListNames.lst_approvalHistory);
            if (isSMERole) {
              this._setCurrentUserRole(smeApprovalHistory, "SME");
            }
          }
          if (generateOtherFunctionsHistory) {
            let otherFunApprovalHistory = approvalHistory.filter(his => { return (his.Role === 'Other Functions') });
            if (otherFunApprovalHistory.length !== 0) {
              this.setState({ isShowOtherFunctionsTab: true });
            }
            this._generateApprovalHistory('OtherFunctions', otherFunApprovalHistory, spListNames.lst_approvalHistory);
            if (isOtherFunctionsRole) {
              this._setCurrentUserRole(otherFunApprovalHistory, "OtherFunctions");
            }
          }
          if (generateCXHistory) {
            let cxApprovalHistory = approvalHistory.filter(his => { return (his.Role === 'CX') });
            if (cxApprovalHistory.length !== 0) {
              this.setState({ isShowCXTab: true });
            }
            this._generateApprovalHistory('CX', cxApprovalHistory, spListNames.lst_approvalHistory);
            if (isCXRole) {
              this._setCurrentUserRole(cxApprovalHistory, "CX");
            }
          }
        }).catch((error) => { });
    }
  }
  public _setCurrentUserRole(historyItems, stage) {
    //Setting current user role basis approval history. Checks if user is approver and approval is done or not
    let isUserApprover = false;
    let currentApprovalItem = 0;
    let isApprovalDone = false;
    historyItems.forEach(historyItem => {
      if (this.state.currentUser.Id === historyItem.PendingWithId && historyItem.State !== 'Completed') {
        isUserApprover = true;
        historyItem.ReviewComments = historyItem.ReviewComments === null ? "" : historyItem.ReviewComments;
        currentApprovalItem = historyItem
        if (historyItem.Status === "Approved") {
          isApprovalDone = true;
          this.setState({});
        }
      }
    });
    //If approver setting user role
    if (isUserApprover) {
      this.setState({ approvalItem: currentApprovalItem, approvalDone: isApprovalDone });
      switch (stage) {
        case 'SME': this.setState({ isSMELoggedIn: true });
          break;
        case 'OtherFunctions': this.setState({ isOtherFunctionApproverLoggedIn: true });
          break;
        case 'CX': this.setState({ isCXApproverLoggedIn: true });
          break;
      }
    }
    this.setState({ isLoading: false });
  }
  //File change event
  public _onAttachmentFileChange = event => {
    this.setState({ selectedAttachmentFile: event.target.files[0] });
  }
  //Upload file to library
  public async _onAttachmentFileUpload(): Promise<void> {
    try {
      if (this.importFileUploadRef.current.files.length === 0) {
        alert("Please select a file");
        return;
      }
      this.setState({ isShowLoader: true })
      //File path is predefined format Document number/revision number
      let filePath = spListNames.lst_initiativeDocuments + "/" + this.state.documentNumber + "/Revision " + this.state.revision + "/";
      let uploadedFiles = await this._Provider.updateDocument(this.importFileUploadRef.current.files[0], this.state.globalID, filePath);
      this.importFileUploadRef.current.value = "";
      //Setting doc preview once file is uploaded
      let docPreview = await this._setDocPreview().then(() => {
        this.setState({ isShowLoader: false });
      })
        .catch((error) => {
          alert("Something went wrong. Please try again later");
          console.log(error);
          window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
        });
    }
    catch (error) {
      this._handlingException(error);
    }
  }
  public _onControlChange = (ev: any, newValue: any): void => {
    //Capturing field control change events. Identifying controls basis field id.
    let controlId = ev.target.id;
    switch (controlId) {
      case "title": this.setState({ title: newValue.toString() !== "" ? newValue.toString() : null });
        break;
      case "comments": this.setState({ comments: newValue.toString() !== "" ? newValue.toString() : null });
        break;
      case "description": this.setState({ description: newValue.toString() !== "" ? newValue.toString() : null });
        break;
      case "documentType": let item = newValue as IDropdownOption;
        let value = parseInt(item.key.toString());
        this.setState({ selectedDocumentType: value }, () => {
          //callback
          console.log(this.state.selectedDocumentType);
          this.generateInitiativeOtherFunctionApprovers().then(() => { }).catch((error) => { console.log(error); });
        });
        this.setState({ isShowLoader: true });

        break;
      case "department": let itemdept = newValue as IDropdownOption;
        let deptvalue = parseInt(itemdept.key.toString());
        this.setState({ selectedDepartment: deptvalue });
        break;
      case "idco_val":
        let selectedIDCOItem = newValue as IDropdownOption;
        if (!selectedIDCOItem) {
          return null;
        }
        let selectedIDCOVal = this.state.allIDCOOptions.filter(
          (idco) => idco.ID === selectedIDCOItem.key
        );
        let selectedIDCO = selectedIDCOVal[0];
        let OFApprover = this.state.selectedOFApproverGP;
        let newObj = { function: OFApprover.function, functionId: OFApprover.functionId, docTypeID: OFApprover.docTypeID, ApproverId: selectedIDCO.ApproverId, ApproverEmail: selectedIDCO.Approver.EMail, itemId: OFApprover.itemId }
        this.setState({ selectedOFApproverGP: newObj });
        this.setState({
          selectedIDCO: parseInt(selectedIDCOItem.key.toString()),
          selectedIDCOObj: { Description: selectedIDCO.Description, ApproverId: selectedIDCO.ApproverId, ApproverEmail: selectedIDCO.Approver.EMail }
        });
        break;
      case "isIDCO":
        let itemidco = newValue as IDropdownOption;
        let valueidco = itemidco.key.toString();
        this.setState({ is_idco: valueidco });
        if (valueidco === 'No') {
          this.setState({
            selectedIDCO: null,
            selectedIDCOObj: null
          });
        }
        break;
      case "reviewComments": this.setState({ reviewComments: newValue.toString() !== "" ? newValue.toString() : null });
        break;
      case "documentNumber": this.setState({ documentNumber: newValue.toString() !== "" ? newValue.toString() : null });
        break;
      case "revision": this.setState({ revision: newValue.toString() !== "" ? newValue.toString().trim() : null });
        break;
      case "Requestforexpiration":
        this.setState({ Requestforexpiration: !this.state.Requestforexpiration })
        break;
      case "isSMEEvaluationRequired":
        let issmerequired = newValue as IDropdownOption;
        let valuesmerequired = issmerequired.key.toString();
        this.setState({ isSMEEvaluationRequired: valuesmerequired });
        if (valuesmerequired === 'No') {
          this.setState({
            selectedSME: [],
            defaultSME: [],
          });
        }
        break;
      case "discussionComment": this.setState({ discussionComment: newValue.toString() !== "" ? newValue.toString() : null });
        break;

      default: break;
    }
  }
  //Capturing SME people picker change event
  public _onSMEChange = (items: any): void => {
    this.setState({ selectedSME: items });
    // if (items.length > 0) {
    //   items.map(item => {
    //     if (item.loginName !== null && item.loginName !== undefined) {
    //       let newObj={ Id: item.id, EMail: item.loginName.split("|membership|")[1] };
    //       this.setState(prevState => ({
    //         selectedSME: [...prevState.selectedSME, newObj]
    //       }))
    //       //this.setState({ selectedSME: { Id: item.id, EMail: item.loginName.split("|membership|")[1] } });
    //     }
    //   });
    //   console.log(this.state.selectedSME)
    // }
    // else {
    //   this.setState({ selectedSME: null });
    // }

  }
  public _onRequesterChange = (items: any): void => {
    if (items.length > 0) {
      items.map(item => {
        if (item.loginName !== null && item.loginName !== undefined) {
          this.setState({ selectedRequester: { RequesterId: item.id, RequesterEmail: item.loginName.split("|membership|")[1], RequesterTitle: item.text } });
        }
      });
    }
    else {
      this.setState({ selectedRequester: null });
    }

  }
  public _generateApprovalHistory(section, approvalHistory, listName) {
    //Generating list data basis approval history items for corresponding tabs
    let items = [];
    let fun = '';
    //fun value will be different basis the approval level. SMEs does not have function value.
    //Loop through each item and defining details list array.
    approvalHistory.forEach(historyItem => {
      switch (section) {
        case 'SME Approval':
        case 'SME Evaluation': fun = '';
          break;
        case 'OtherFunctions': fun = (historyItem.OtherFunctions === undefined || historyItem.OtherFunctions === null) ? "" : historyItem.OtherFunctions.Title;
          break;
        case 'CX': fun = (historyItem.CXFunction === undefined || historyItem.CXFunction === null) ? "" : historyItem.CXFunction.Title;
          break;
        default: fun = '';
      }
      items.push({
        key: historyItem.ID,
        function: fun,
        actionedDate: moment(historyItem.ActionedDate).format('DD MMM YYYY hh:mm A'),
        status: historyItem.Status,
        comments: historyItem.ReviewComments,
        PendingWith: historyItem.PendingWith,
        action: "",
        listName: listName,
        state: historyItem.State,
        role: historyItem.Role
      });
    });
    let columns: IColumn[] = [
      {
        key: 'Approver',
        name: 'Approver',
        fieldName: 'approver',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'Function',
        name: 'Function',
        fieldName: 'function',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'Status',
        name: 'Status',
        fieldName: 'status',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'ReviewComments',
        name: 'Review Comments',
        fieldName: 'comments',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: 'Action',
        name: 'Action',
        fieldName: 'action',
        isMultiline: true,
        minWidth: 200,
        isResizable: true,
      }
    ];
    //Filtering details list columns basis function value and approval status
    switch (section) {
      case 'SME Evaluation':
        columns = columns.filter(col => { return (col.key !== 'Function') });
        if (!this.state.isUnderSMEEvaluation)
          columns = columns.filter(col => { return (col.key !== 'Action') });
        this.setState({ approvalHistorySMEEvaluationColumns: columns });
        this.setState({ approvalHistorySMEEvaluation: items });
        break;
      case 'SME Approval':
        columns = columns.filter(col => { return (col.key !== 'Function') });
        if (!this.state.isInApprovalSME)
          columns = columns.filter(col => { return (col.key !== 'Action') });
        this.setState({ approvalHistorySMEColumns: columns });
        this.setState({ approvalHistorySME: items });
        break;
      case 'OtherFunctions':
        if (!this.state.isInApprovalOtherFunction)
          columns = columns.filter(col => { return (col.key !== 'Action') });
        this.setState({ approvalHistoryOtherFunctionsColumns: columns });
        this.setState({ approvalHistoryOtherFunctions: items });
        break;
      case 'CX':
        if (!this.state.isInApprovalCX)
          columns = columns.filter(col => { return (col.key !== 'Action') });
        this.setState({ approvalHistoryCXColumns: columns });
        this.setState({ approvalHistoryCX: items }); break;
    }
  }
  public async _sendForApproval(id, listName, comments) {
    //Sending for approval if rejected or sent back to requester
    this.setState({ isShowLoader: true });
    let reviewComment = "<p><span>" + moment().format('DD MMM YYYY hh:mm A') + " (In review)</span></p>" + comments;
    let updateValue = {
      Status: 'In review',
      ActionedDate: moment(),
      ReviewComments: reviewComment,
      TriggerWF: true
    };
    await this._Provider.updateItem(listName, updateValue, id)
      .then((result) => {
        // window.location.reload()
      }).catch((error) => {
        console.log(error);
      });
    //We are using custom column LastUpdatedAt to set modifiedby inorder to set correct modified by user as we are using differnt list for approval history
    let updateInitiativeValue = {
      LastUpdatedAt: moment(),
      LastUpdatedById: this.state.currentUser.Id
    };
    let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateInitiativeValue, this.state.globalID).then((result) => {
      alert("Successfully sent for approval")
      window.location.reload()
    }).catch((error) => {
      console.log(error);
    });
    //window.location.href = this.props.context.pageContext.web.absoluteUrl;
  }
  public async _onApproverAction(section, action) {
    //Updating approver action according to approver level and action
    let actioneText = action;
    this.setState({ isShowLoader: true });
    let reviewComments = "", listName = "", id = 0;
    let actionedDate = "<span>" + moment().format('DD MMM YYYY hh:mm A') + "</span>";
    reviewComments = this.state.approvalItem.ReviewComments;
    listName = spListNames.lst_approvalHistory;
    id = this.state.approvalItem.Id;
    if (section === "SME Evaluation" && action === "Approved") {
      actioneText = "Completed";
    }
    if (this.state.reviewComments !== null)
      reviewComments = "<p><span>" + actionedDate + " (" + actioneText + ")</span></br>" + this.state.reviewComments + "</p>" + reviewComments;
    else
      reviewComments = "<p><span>" + actionedDate + " (" + actioneText + ")</span></p>" + reviewComments;
    let updateValue = {
      Status: action,
      ActionedDate: moment(),
      ReviewComments: reviewComments,
      TriggerWF: true
    };
    await this._Provider.updateItem(listName, updateValue, id)
      .then((result) => {
        //alert("Successfully updated")
        // window.location.reload()
      }).catch((error) => {
        console.log(error);
      });
    //We are using custom column LastUpdatedAt to set modifiedby inorder to set correct modified by user as we are using differnt list for approval history
    let updateInitiativeValue = {
      LastUpdatedAt: moment(),
      LastUpdatedById: this.state.currentUser.Id
    };
    let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateInitiativeValue, this.state.globalID).then((result) => {
      alert("Successfully updated");
      window.location.reload();
    }).catch((error) => {
      console.log(error);
    });

  }
  public async _onPostDiscussion() {
    //Updating approver action according to approver level and action
    this.setState({ isShowLoader: true });

    let insertValue = {
      Comment: this.state.discussionComment,
      InitiativeIdId: this.state.globalID
    };
    await this._Provider.createSPListItem(insertValue, spListNames.lst_initiativeDiscussions)
      .then((result) => {
        //alert("Successfully updated")
        // window.location.reload()
      }).catch((error) => {
        console.log(error);
      });
    //We are using custom column LastUpdatedAt to set modifiedby inorder to set correct modified by user as we are using differnt list for approval history
    let updateInitiativeValue = {
      LastUpdatedAt: moment(),
      LastUpdatedById: this.state.currentUser.Id
    };
    let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateInitiativeValue, this.state.globalID).then((result) => {
      alert("Successfully posted");

    }).catch((error) => {
      window.location.reload();
      console.log(error);
    });
    let [discussions] = await Promise.all([this._Provider.getSPListItems("InitiativeIdId eq  " + this.state.globalID, 'Created', false, '*,Author/Id,Author/EMail,Author/Name,Author/Title', 'Author', spListNames.lst_initiativeDiscussions)])
    //Setting form control inputs
    this.setState({
      discussionComment: "",
      discussions: discussions,
      isShowLoader: false
    });

  }
  //This function is part of details list
  private _getKey(item: any, index?: number): string {
    return item.key;
  }
  public onRenderItemColumn = (item, index: number, column: IColumn): JSX.Element | React.ReactText => {
    //Constructing html outputs to render list columns
    //Review comments will be rendered as hover card
    //Send for approval will be enabled for requester if rejected or sent back to requester by approver.
    const onRenderPlainCard = (item: any): JSX.Element => {
      return (
        <div className={itemClasses.hoverItem} dangerouslySetInnerHTML={{ __html: item.comments }}>
        </div>
      );
    };
    const plainCardProps: IPlainCardProps = {
      onRenderPlainCard: onRenderPlainCard,
      renderData: item,
    };
    switch (column.key) {
      case "ReviewComments":
        return (
          <HoverCard plainCardProps={plainCardProps} instantOpenOnClick type={HoverCardType.plain} className={itemClasses.cards}>
            <div className={itemClasses.label}>
              View
            </div>
          </HoverCard>
        );
      case "Approver":
        return (
          <div>
            <img className={mergeStyles({ width: '25px', borderRadius: '50px', marginRight: '10px' })} src={this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&accountname=" + item.PendingWith.EMail}></img>
            {item.PendingWith.Title}
          </div>
        );
      case "Status":
        let actionedHtml = <p className={mergeStyles({ marginBottom: '0px' })}><span className={mergeStyles({ color: '#808080', fontSize: "11px" })}>{item.actionedDate}</span></p>;
        switch (item.status) {
          case 'Approved':
            let statusApproveVal = item.status;
            if (item.role === 'SME Evaluation') {
              statusApproveVal = "Completed";
            }
            return <><p className={mergeStyles({ marginBottom: '0px' })}><span className={mergeStyles({ color: 'green', height: '100%', display: 'block' })}>{statusApproveVal}</span></p>{actionedHtml}</>;
          case 'In review': return <><p className={mergeStyles({ marginBottom: '0px' })}><span className={mergeStyles({ height: '100%', display: 'block' })}>{item.status}</span></p>{actionedHtml}</>;
          case 'Sent back to requester':
            return <><p className={mergeStyles({ marginBottom: '0px' })}><span className={mergeStyles({ color: '#f77301', height: '100%', display: 'block' })}>{item.status}</span></p>{actionedHtml}</>;
          case "Rejected": return <><p className={mergeStyles({ marginBottom: '0px' })}><span className={mergeStyles({ color: '#A4262C', height: '100%', display: 'block' })}>{item.status}</span></p>{actionedHtml}</>;
        }
        return <></>;
      case "Action":
        let actionButtonText = "Send for approval";
        if (item.role === 'SME Evaluation') {
          actionButtonText = "Send for evaluation";
        }
        if (item.status !== 'Approved' && item.status !== 'In review' && item.state !== 'Completed' && this.state.isAuthorLoggedIn)
          return <span> <CommandBarButton style={{ display: 'inline-block', padding: '5px 2px', fontSize: "12px" }} className={style.button} primary={true} text={actionButtonText} onClick={() => this._sendForApproval(item.key, item.listName, item.comments).then(() => { }).catch(() => { })} /></span>;
        else if (item.state === 'Completed')
          return <>Approval flow completed</>
        else
          return <></>
      case "Function": return <span>{item.function}</span>;
    }

    return item[column.key];
  };
  public async _onsubmit(eventType): Promise<void> {
    //On submit button click function ensure mandatory fields and update the initiative list accordingly
    try {
      let actionedDate = "<p><span>" + moment().format('DD MMM YYYY hh:mm A') + " (In review)</span></p>";
      let validation = this._formValidation(eventType);
      let updateValue = null;
      if (validation === "") {
        this.setState({ isShowLoader: true });
        let smeResultObj = [];
        if (this.state.selectedSME.length > 0) {
          this.state.selectedSME.map(item => {
            if (item.id !== null && item.id !== undefined) {
              smeResultObj.push(item.id)
            }
          });
        }
        if (eventType === "AdminUpdate") {
          let filter = "DocumentNumber eq '" + this.state.documentNumber + "' and Revision eq '" + this.state.revision + "'";
          let SPItems = await this._Provider.getSPListItems(filter, 'ID', true, '*', '', spListNames.lst_initiatives,);
          let delegated = null;
          if (this.state.createdBy !== this.state.selectedRequester.RequesterTitle) {
            delegated = moment();
          }
          if (SPItems.length === 0) {
            updateValue = {
              Title: this.state.title,
              Description: this.state.description,
              Status: "Draft",
              Revision: this.state.revision,
              DocumentNumber: this.state.documentNumber,
              InitiateItem_UpdateWF: true,
              DocumentTypeId: this.state.selectedDocumentType,
              LastUpdatedAt: moment(),
              LastUpdatedById: this.state.currentUser.Id,
              RequesterId: this.state.selectedRequester.RequesterId,
              PreviousRequesterId: this.state.previousRequesterId,
              Delegatedat: delegated
            };
            let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID)
            let OFApproverInsert = await this.generatePushOFApproversObject("Update");
            alert("Record has been updated");
            window.location.href = this.props.context.pageContext.web.absoluteUrl + '/SitePages/My-Documents.aspx';
          }
          else {
            alert("Duplicate record found for this Document number - Revision combination");
            this.setState({ isShowLoader: false });
          }

        }
        else if (eventType === "ChangeRequester") {
          updateValue = {
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id,
            RequesterId: this.state.selectedRequester.RequesterId,
            InitiateChangeRequesterWF: true,
            Delegatedat: moment(),
            //InitiateItem_UpdateWF: true,
            PreviousRequesterId: this.state.previousRequesterId
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID)
          alert("Requester has been updated");
          window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
        }
        else if (eventType === 'NewRequest') {
          let newItem = null;
          newItem = {
            Status: "Requested",
            RequestorComments: this.state.comments,
            Title: this.state.title,
            Description: this.state.description,
            DocumentTypeId: this.state.selectedDocumentType,
            InitiateItem_UpdateWF: true,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id,
            RequesterId: this.state.currentUser.Id,
            PreviousRequesterId: this.state.currentUser.Id,
            CreatedByUser: this.state.currentUser.Title
          };
          let result = await this._Provider.createItem(newItem);
          this.setState({ globalID: result.data.ID }, () => {
            this.generatePushOFApproversObject("Create").then(() => {
              alert("Request for new initiative is submitted successfully \n You will be notified ");
              window.location.href = this.props.context.pageContext.web.absoluteUrl;
            }).catch((error) => { alert("Something went wrong. Please try again later"); });
          });
        }
        else if (eventType === 'NewRevisionRequest') {
          let filter = "DocumentNumber eq '" + this.state.documentNumber + "' and Revision eq '" + this.state.revision + "'";
          let SPItems = await this._Provider.getSPListItems(filter, 'ID', true, '*', '', spListNames.lst_initiatives,);
          if (SPItems.length === 0) {
            let newItem = null;
            newItem = {
              Status: "Draft",
              Title: this.state.title,
              Description: this.state.description,
              DocumentNumber: this.state.documentNumber,
              DocumentTypeId: this.state.selectedDocumentType,
              DepartmentId: this.state.selectedDepartment,
              Revision: this.state.revision,
              InitiateItem_UpdateWF: true,
              Requestforexpiration: this.state.Requestforexpiration,
              LastUpdatedAt: moment(),
              LastUpdatedById: this.state.currentUser.Id,
              RequesterId: this.state.currentUser.Id,
              PreviousRequesterId: this.state.currentUser.Id,
              CreatedByUser: this.state.currentUser.Title
            };
            let result = await this._Provider.createItem(newItem);
            this.setState({ globalID: result.data.ID }, () => {
              this.generatePushOFApproversObject("Create").then(() => {
                alert("Request for new revision is submitted successfully. ");
                window.location.href = this.props.context.pageContext.web.absoluteUrl + '/SitePages/My-Documents.aspx';
              }).catch((error) => { alert("Something went wrong. Please try again later"); });
            });

          }
          else {
            alert("Duplicate record found for this Document number - Revision combination");
          }
        }
        if (eventType === 'EditRequest-Draft') {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Draft",
            DocumentTypeId: this.state.selectedDocumentType,
            DepartmentId: this.state.selectedDepartment,
            IDCO_x0028_Yes_x002f_No_x0029_: this.state.is_idco === "Yes" ? true : false,
            IDCOId: this.state.selectedIDCO,
            IDCODescription: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.Description : null,
            IDCOApproverId: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.ApproverId : null,
            SMEId: this.state.selectedSME.length > 0 !== null ? { 'results': smeResultObj } : null,
            SMEEvaluationRequired: this.state.isSMEEvaluationRequired === "Yes" ? true : false,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID)
          let OFApproverInsert = await this.generatePushOFApproversObject("Update");
          alert("Record has been updated");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === 'Send For SME Evaluation') {
          if (this.state.initiativeDocumentArray.length === 0) {
            alert("Please upload document");
            this.setState({ isShowLoader: false });
            return;
          }
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Under SME Evaluation",
            DocumentTypeId: this.state.selectedDocumentType,
            DepartmentId: this.state.selectedDepartment,
            IDCO_x0028_Yes_x002f_No_x0029_: this.state.is_idco === "Yes" ? true : false,
            IDCOId: this.state.selectedIDCO,
            IDCODescription: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.Description : null,
            IDCOApproverId: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.ApproverId : null,
            SMEId: this.state.selectedSME !== null ? { 'results': smeResultObj } : null,
            InitiateItem_UpdateWF: true,
            ItemActionedDate: actionedDate,
            SMEEvaluationRequired: this.state.isSMEEvaluationRequired === "Yes" ? true : false,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID)
          let OFApproverInsert = await this.generatePushOFApproversObject("Update");
          alert("SME evaluation initiated");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === 'Send For SME Approval') {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Under Approval (SME)",
            DocumentTypeId: this.state.selectedDocumentType,
            DepartmentId: this.state.selectedDepartment,
            IDCO_x0028_Yes_x002f_No_x0029_: this.state.is_idco === "Yes" ? true : false,
            IDCOId: this.state.selectedIDCO,
            IDCODescription: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.Description : null,
            IDCOApproverId: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.ApproverId : null,
            SMEId: this.state.selectedSME !== null ? { 'results': smeResultObj } : null,
            InitiateItem_UpdateWF: true,
            ItemActionedDate: actionedDate,
            SMEEvaluationRequired: this.state.isSMEEvaluationRequired === "Yes" ? true : false,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          let OFApproverInsert = await this.generatePushOFApproversObject("Update");
          alert("Successfully Sent for SME Approval");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === 'Send for Other Functions Approval') {
          if (this.state.initiativeDocumentArray.length === 0) {
            alert("Please upload document");
            this.setState({ isShowLoader: false });
            return;
          }
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Under Approval (Other Functions)",
            InitiateItem_UpdateWF: true,
            ItemActionedDate: actionedDate,
            IDCO_x0028_Yes_x002f_No_x0029_: this.state.is_idco === "Yes" ? true : false,
            IDCOId: this.state.selectedIDCO,
            IDCODescription: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.Description : null,
            IDCOApproverId: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.ApproverId : null,
            SMEEvaluationRequired: this.state.isSMEEvaluationRequired === "Yes" ? true : false,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id,
            OtherFunctionApprovalRequired:true
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          let OFApproverInsert = await this.generatePushOFApproversObject("Update");
          alert("Successfully Sent for Other Functions Approval");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === 'Send for CX Approval') {
          if(this.state.isDraftStage && this.state.isSMEEvaluationRequired === "No"){
            updateValue = {
              Title: this.state.title,
              Description: this.state.description,
              Status: "Under Approval (CX)",
              DocumentTypeId: this.state.selectedDocumentType,
              DepartmentId: this.state.selectedDepartment,
              IDCO_x0028_Yes_x002f_No_x0029_: this.state.is_idco === "Yes" ? true : false,
              IDCOId: this.state.selectedIDCO,
              IDCODescription: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.Description : null,
              IDCOApproverId: this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.ApproverId : null,
              SMEId: this.state.selectedSME !== null ? { 'results': smeResultObj } : null,
              InitiateItem_UpdateWF: true,
              ItemActionedDate: actionedDate,
              LastUpdatedAt: moment(),
              LastUpdatedById: this.state.currentUser.Id,
              OtherFunctionApprovalRequired:false,
              SMEEvaluationRequired: this.state.isSMEEvaluationRequired === "Yes" ? true : false
            };
          }
          else{
            updateValue = {
              Title: this.state.title,
              Description: this.state.description,
              Status: "Under Approval (CX)",
              InitiateItem_UpdateWF: true,
              ItemActionedDate: actionedDate,
              LastUpdatedAt: moment(),
              LastUpdatedById: this.state.currentUser.Id
            };
          }
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Sent for CX Approval");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === "Send for Closure / Approved") {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Sent for Closure / Approved",
            ItemActionedDate: actionedDate,
            InitiateItem_UpdateWF: true,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Sent for Closure/Approved");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === "Send for Closure / Rejected") {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Sent for Closure / Rejected",
            InitiateItem_UpdateWF: true,
            ItemActionedDate: actionedDate,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Sent for Closure/Rejected");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === "Send for Closure / Expired") {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Sent for Closure / Expired",
            InitiateItem_UpdateWF: true,
            ItemActionedDate: actionedDate,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Sent for Closure/Expired");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === "Close / Approved") {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Closed / Approved",
            ItemActionedDate: actionedDate,
            InitiateItem_UpdateWF: true,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Closed/Approved");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === "Close / Rejected") {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Closed / Rejected",
            ItemActionedDate: actionedDate,
            InitiateItem_UpdateWF: true,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Closed/Rejected");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === "Close / Expired") {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Closed / Expired",
            ItemActionedDate: actionedDate,
            InitiateItem_UpdateWF: true,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Closed/Expired");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
        if (eventType === "Close / Cancelled") {
          updateValue = {
            Title: this.state.title,
            Description: this.state.description,
            Status: "Closed / Cancelled",
            ItemActionedDate: actionedDate,
            InitiateItem_UpdateWF: true,
            LastUpdatedAt: moment(),
            LastUpdatedById: this.state.currentUser.Id
          };
          let result = await this._Provider.updateItem(spListNames.lst_initiatives, updateValue, this.state.globalID);
          alert("Successfully Closed/Cancelled");
          this._getInitiativeDetails().then(() => {
            this.setState({ isShowLoader: false });
          }).catch((error) => {
            alert("Something went wrong. Please try again later");
            console.log(error);
            window.location.href = window.location.href.split("?")[0] + "?DocID=" + this.state.globalID;
          });
        }
      }
      else {
        alert("Please fill out below mandatory fields \n" + validation);
      }
    }
    catch (error) {
      this._handlingException(error);
    }
  }
  public _formValidation(stage) {
    //Checks for mandatory fields
    let temp = "";
    if (stage === 'AdminUpdate') {
      if (this.state.selectedDocumentType === null) {
        temp = temp + "Document Type\n";
      }
      if (this.state.title === null) {
        temp = temp + "Title\n";
      }
      if (this.state.description === null) {
        temp = temp + "Description\n";
      }
      if (this.state.documentNumber === null) {
        temp = temp + "Document Number\n";
      }
      if (this.state.revision === null) {
        temp = temp + "Revision\n";
      }
      if (this.state.selectedRequester === null) {
        temp = temp + "Requester\n";
      }
    }
    if (stage === 'ChangeRequester') {
      if (this.state.selectedRequester === null) {
        temp = temp + "Requester\n";
      }
    }
    if (stage === "NewRevisionRequest") {
      if (this.state.selectedDocumentType === null) {
        temp = temp + "Document Type\n";
      }
      if (this.state.title === null) {
        temp = temp + "Title\n";
      }
      if (this.state.description === null) {
        temp = temp + "Description\n";
      }
      if (this.state.documentNumber === null) {
        temp = temp + "Document Number\n";
      }
      if (this.state.revision === null) {
        temp = temp + "Revision\n";
      }
    }
    else if (stage === 'NewRequest') {
      if (this.state.comments === null) {
        temp = temp + "Comments\n";
      }
      if (this.state.selectedDocumentType === null) {
        temp = temp + "Document Type\n";
      }
      if (this.state.title === null) {
        temp = temp + "Title\n";
      }
      if (this.state.description === null) {
        temp = temp + "Description\n";
      }
    }
    else if (stage === "EditRequest-Draft") {
      if (this.state.title === null) {
        temp = temp + "Title\n";
      }
      if (this.state.description === null) {
        temp = temp + "Description\n";
      }
      if (this.state.selectedDocumentType === null || this.state.selectedDocumentType === undefined) {
        temp = temp + "Document Type\n";
      }
      if (this.state.selectedDepartment === null || this.state.selectedDepartment === undefined) {
        temp = temp + "Department\n";
      }
      if (this.state.is_idco === "Yes") {
        if (this.state.selectedIDCO === null || this.state.selectedIDCO === undefined) {
          temp = temp + "IDCO\n";
        }
      }
      if (this.state.isSMEEvaluationRequired === "Yes" && this.state.selectedSME.length === 0) {
        temp = temp + "SME\n";
      }
      if (this.state.selectedOFApproverGQ.ApproverId === null) {
        temp = temp + "Other Function Approver - Global Quality\n";
      }
      if (this.state.selectedOFApproverGME.ApproverId === null) {
        temp = temp + "Other Function Approver - Global Manufacturing Engineering\n";
      }
      if (this.state.selectedOFApproverGP.ApproverId === null) {
        temp = temp + "Other Function Approver - Global Purchasing \n";
      }
    }
    else if (stage === "Send For SME Approval" || stage === "Send For Other Functions Approval" || stage === "Send For SME Evaluation") {
      if (this.state.title === null) {
        temp = temp + "Title\n";
      }
      if (this.state.description === null) {
        temp = temp + "Description\n";
      }
      if (this.state.selectedDocumentType === null) {
        temp = temp + "Document Type\n";
      }
      if (this.state.is_idco === "Yes") {
        if (this.state.selectedIDCO === null || this.state.selectedIDCO === undefined) {
          temp = temp + "IDCO\n";
        }
      }
      if (this.state.isSMEEvaluationRequired === "Yes" && this.state.selectedSME.length === 0) {
        temp = temp + "SME\n";
      }
      if (this.state.selectedOFApproverGQ.ApproverId === null) {
        temp = temp + "Other Function Approver - Global Quality\n";
      }
      if (this.state.selectedOFApproverGME.ApproverId === null) {
        temp = temp + "Other Function Approver - Global Manufacturing Engineering\n";
      }
      if (this.state.selectedOFApproverGP.ApproverId === null) {
        temp = temp + "Other Function Approver - Global Purchasing \n";
      }

    }
    else if (stage === "Send For CX Approval") {
      if (this.state.title === null) {
        temp = temp + "Title\n";
      }
      if (this.state.description === null) {
        temp = temp + "Description\n";
      }
      if (this.state.selectedDocumentType === null) {
        temp = temp + "Document Type\n";
      }
      if (this.state.is_idco === "Yes") {
        if (this.state.selectedIDCO === null || this.state.selectedIDCO === undefined) {
          temp = temp + "IDCO\n";
        }
      }
      if (this.state.isSMEEvaluationRequired === "Yes" && this.state.selectedSME.length === 0) {
        temp = temp + "SME\n";
      }

    }

    return temp;
  }
  public _handlingException(error) {
    alert("There is an error\n" + error);
    console.log(error);
  }
  public render(): React.ReactElement<ITechnicalSpecificationInputFormProps> {
    const {
      userDisplayName,
      context,
      Title,
    } = this.props;
    //Setting form template basis request status
    return (
      <div>
        {
          this.state.isNewStage && (
            <div>
              {this.renderNewForm()}
            </div>
          ) ||
          this.state.isNewRevStage && (
            <div>
              {this.renderNewRevisionForm()}
            </div>
          ) ||
          (this.state.isRequestedStage && this.state.isAdminLoggedIn) && (
            <div>
              {this.renderAdminForm()}
            </div>
          ) ||
          (!this.state.isNewStage && !this.state.isRequestedStage) && (
            !this.state.isLoading && (
              <div>
                {this.renderEditForm()}
              </div>
            ) ||
            this.state.isLoading && (
              <div>
                {this.rednderLoader()}
              </div>
            )
          ) ||
          this.state.isNotAuthorised && (
            <div>
              {this.renderNotAuthorisedForm()}
            </div>
          )
        }
      </div >
    );
  }
  public renderNewForm = (): JSX.Element => {
    //New form will be rendered
    return (
      <Fabric className={style.boxShadow}>
        <div className={style.gridContainer} style={{ padding: '20px' }}>
          {/* <div className="col-sm-12" style={{ textAlign: "right" }}><span className='p-2'>Welcome, <b>{this.state.currentUser !== null ? this.state.currentUser.Title : <b></b>}</b></span></div> */}
          <div className={style.topbanner + " row"}>
            <div className="col-sm-12">
              <div className={style.titleRow}>Request New Initiative</div>
            </div>
          </div>
          <div className='row'>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <Dropdown
                id="documentType"
                required={true}
                label="Document Type"
                selectedKey={this.state.selectedDocumentType !== null ? this.state.selectedDocumentType : undefined}
                onChange={this._onControlChange}
                placeholder=""
                options={this.state.dropdownDocumentTypeOptions}
                styles={dropdownStyles}
              />
            </div>
            <div className="col-sm-12">
              <TextField label="Title" id="title" defaultValue={this.state.title} value={this.state.title} placeholder="" onChange={this._onControlChange} styles={textFieldStyles} required />
            </div>
            <div className="col-sm-12">
              <TextField label="Description" id="description" defaultValue={this.state.description} placeholder="" value={this.state.description} multiline rows={2} onChange={this._onControlChange} styles={textFieldStyles} required />
            </div>
            <div className={style.fieldCols + " col-sm-12"}>
              <TextField label="Requester Comments" id="comments" placeholder="" defaultValue={this.state.comments} value={this.state.comments} multiline rows={2} onChange={this._onControlChange} styles={textFieldStyles} />
            </div>
            <div className="col-lg-12 col-md-12 col-sm-12" style={{ textAlign: "center", paddingTop: "10px" }}>
              <CommandBarButton style={{ display: 'inline-block' }} className={style.button} styles={buttonStyles} primary={true} text="Submit" onClick={() => { this._onsubmit("NewRequest").then(() => { }).catch(() => { }); }} />
            </div>
          </div>
        </div>
      </Fabric>
    )
  }
  public renderNewRevisionForm = (): JSX.Element => {
    //Render revision request form
    return (
      <Fabric className={style.boxShadow}>
        <div className={style.gridContainer} style={{ padding: '20px' }}>
          <div className={style.topbanner + " row"}>
            <div className="col-sm-12">
              <div className={style.titleRow}>Request New Revision</div>
            </div>
          </div>
          <div className='row'>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <TextField id="documentNumber" label="Document Number" defaultValue={this.state.documentNumber} placeholder="" value={this.state.documentNumber} onChange={this._onControlChange} styles={textFieldStyles} required={true} />
            </div>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"} >
              <TextField id="revision" label="Revision" defaultValue={this.state.revision} placeholder="" value={this.state.revision} onChange={this._onControlChange} styles={textFieldStyles} required={true} />
            </div>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <Checkbox id="Requestforexpiration" label="Request for expiration" checked={this.state.Requestforexpiration} onChange={this._onControlChange} />
            </div>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <Dropdown
                id="documentType"
                required={true}
                label="Document Type"
                selectedKey={this.state.selectedDocumentType !== null ? this.state.selectedDocumentType : undefined}
                onChange={this._onControlChange}
                placeholder=""
                options={this.state.dropdownDocumentTypeOptions}
                styles={dropdownStyles}
              />
            </div>
            <div className="col-sm-12">
              <TextField label="Title" id="title" defaultValue={this.state.title} value={this.state.title} onChange={this._onControlChange} placeholder="" styles={textFieldStyles} required />
            </div>
            <div className="col-sm-12">
              <TextField label="Description" id="description" defaultValue={this.state.description} placeholder="" value={this.state.description} multiline rows={2} onChange={this._onControlChange} styles={textFieldStyles} required />
            </div>
            <div className="col-lg-12 col-md-12 col-sm-12" style={{ textAlign: "center", paddingTop: "10px" }}>
              <CommandBarButton style={{ display: 'inline-block' }} className={style.button} styles={buttonStyles} primary={true} text="Submit" onClick={() => { this._onsubmit("NewRevisionRequest").then(() => { }).catch(() => { }); }} />
            </div>
          </div>
        </div>
      </Fabric>
    )
  }
  public renderAdminForm = (): JSX.Element => {
    //Render this form if request status is 'Requested' and admin user is logged in
    return (
      <Fabric className={style.boxShadow}>
        <div className={style.gridContainer} style={{ padding: '20px' }}>
          <div className={style.topbanner + " row"}>
            <div className="col-sm-3"></div>
            <div className="col-sm-6">
              <div className={style.titleRow}>Admin Form</div>
            </div>
          </div>
          <div className='row'>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <PeoplePicker
                context={this.props.context as any}
                titleText="Requested By"
                ensureUser={true}
                personSelectionLimit={1}
                showtooltip={true}
                showHiddenInUI={false}
                onChange={this._onRequesterChange}
                principalTypes={[PrincipalType.User]}
                defaultSelectedUsers={this.state.selectedRequester !== null ? [this.state.selectedRequester.RequesterEmail] : []}
                resolveDelay={1000}
                styles={peoplePickerStyles} />
            </div>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <TextField label="Requester Comments" id="comments" defaultValue={this.state.comments} placeholder="" value={this.state.comments} multiline rows={2} styles={textFieldStyles} />
            </div>

            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <Dropdown
                id="documentType"
                required={true}
                label="Document Type"
                selectedKey={this.state.selectedDocumentType !== null ? this.state.selectedDocumentType : undefined}
                onChange={this._onControlChange}
                placeholder=""
                options={this.state.dropdownDocumentTypeOptions}
                styles={dropdownStyles}
              />
            </div>
            <div className="col-sm-12">
              <TextField label="Title" id="title" defaultValue={this.state.title} value={this.state.title} placeholder="" onChange={this._onControlChange} styles={textFieldStyles} required />
            </div>
            <div className="col-sm-12">
              <TextField label="Description" id="description" defaultValue={this.state.description} value={this.state.description} placeholder="" multiline rows={2} onChange={this._onControlChange} styles={textFieldStyles} required />
            </div>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
              <TextField id="documentNumber" label="Document Number" defaultValue={this.state.documentNumber} placeholder="" styles={textFieldStyles} value={this.state.documentNumber} onChange={this._onControlChange} required={true} />
            </div>
            <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"} >
              <TextField id="revision" label="Revision" defaultValue={this.state.revision} value={this.state.revision} placeholder="" onChange={this._onControlChange} styles={textFieldStyles} required={true} />
              {/* <Dropdown
                id="revision"
                required={true}
                label="Revision"
                selectedKey={this.state.revision !== null ? this.state.revision : undefined}
                onChange={this._onControlChange}
                options={this.state.dropdownRevisions}
                styles={dropdownStyles}
              /> */}
            </div>
            <div className="col-lg-12 col-md-12 col-sm-12" style={{ textAlign: "center", paddingTop: "10px" }}>
              <CommandBarButton style={{ display: 'inline-block' }} className={style.button} styles={buttonStyles} primary={true} text="Submit" onClick={() => { this._onsubmit("AdminUpdate").then(() => { }).catch(() => { }); }} />
            </div>
          </div>
        </div>
      </Fabric>
    )
  }
  public renderEditForm = (): JSX.Element => {
    //This form is rendered if the request is in place.
    return (
      <Fabric className={style.boxShadow} >
        <div className={style.gridContainer}>
          {this.state.isShowLoader && (
            <div className={style.loaderDiv}>
              <Spinner styles={loaderStyles} label="Please wait.." />
            </div>
          )}
          <div className={" row"}>
            {/* Left section is displaying the non editable details and the document preview  */}
            <div className={style.sectionLeft + " col-lg-4"}>
              <div className='row'>
                <div className={"col-sm-6 " + style.leftSecCol}>
                  <p className={style.labelHeader}>Document Number</p>
                  <p className={style.labelContent} style={{ marginBottom: '0px' }}>{this.state.documentNumber}</p>
                </div>
                <div className={"col-sm-6 " + style.leftSecCol}>
                  <p className={style.labelHeader}>Revision</p>
                  <p className={style.labelContent}>{this.state.revision}</p>
                </div>
                {this.state.Requestforexpiration &&
                  (
                    <div className={"col-sm-12 " + style.leftSecCol} style={{ marginTop: '-15px' }}>
                      <p>(Request for expiration)</p>
                    </div>
                  )}
                <div className={"col-12 " + style.leftSecCol}>
                  <p className={style.labelHeader}>Status</p>
                  <p className={style.labelContent}>{this.state.status}</p>
                </div>
              </div>
              <div className={"col-sm-12 " + style.docDetailsCard}>
                <div className={style.sectionDocUserItem}>
                  <Persona
                    {...this.state.requesterPersona}
                    styles={personaStyles}
                  />
                </div>
                <div className={style.sectionDocUserItem}>
                  <Persona
                    {...this.state.editorPersona}
                    styles={personaStyles}
                  />
                </div>
              </div>
              <div className="col-sm-12">
                <div className={style.sectionDocumentThumb}>
                  <p className={style.labelHeader}>Documents</p>
                  <p className={style.labelHeader}>(Click on the preview to access the documents)</p>
                  {this.state.initiativeDocumentArray.map(document => {
                    return (
                      <DocumentCard
                        aria-label=""
                        onClickHref={document.documentInfo.EncodedAbsUrl + "?csf=1&web=1"}
                        onClickTarget="_blank"
                        styles={docCardSTyles}
                      >
                        <DocumentCardPreview {...document.docPreviewProps} styles={docCardImgSTyles} />
                        <DocumentCardTitle
                          title={document.documentInfo.FileLeafRef}
                          shouldTruncate
                          styles={docCardTitleSTyles}
                        />
                        <DocumentCardActivity activity={"Modified at " + document.docModified} people={[document.documentCardActivityPeople]} />
                      </DocumentCard>
                    )
                  })}
                </div>
              </div>
            </div>
            <div className={"col-lg-8 "}>
              <Pivot aria-label="Basic Pivot initiativeNav"
                styles={pivotStyles}>
                {/* Form field details are displayed in first pivot item. Basis the sme evaluation required selection, the sme details field and send for sme evaluation button will be enabled. 
                  Basis IDCO? selection, IDCO details will be enabled. 
                  Button area is included in first pivot. Buttons will be eabled basis the form state and current user access level. 
                  Form fields will be enabled basis form state as well as current user access level.
                  Other pivots displays the approval history. These tabs will be enabled basis the form state as well as it waits for approval tasks to be generated by flows
                  Approver actions are enabled in approval tabs*/}
                <PivotItem
                  headerText="Initiative Details"
                  headerButtonProps={{
                    'data-order': 1,
                    'data-title': 'Initiative Details',
                  }}
                >
                  <div className={ " row " + style.sectionDetails}>
                    <div className='col-12'>
                      {(this.state.isAdminLoggedIn && !this.state.isClosed) && (
                        <div id="adminSection" className={" row"} >
                          <div className={style.fieldCols + " col-lg-6 col-md-6 col-sm-12"}>
                            <PeoplePicker
                              context={this.props.context as any}
                              titleText="Requested By"
                              ensureUser={true}
                              personSelectionLimit={1}
                              showtooltip={true}
                              showHiddenInUI={false}
                              onChange={this._onRequesterChange}
                              principalTypes={[PrincipalType.User]}
                              defaultSelectedUsers={this.state.selectedRequester !== null ? [this.state.selectedRequester.RequesterEmail] : []}
                              resolveDelay={1000}
                              styles={peoplePickerStyles} />
                          </div>
                          <div className={style.fieldCols + " col-lg-6 col-md-6 col-sm-12"} style={{ paddingTop: '25px' }}>
                            <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Update Requester" onClick={() => { this._onsubmit("ChangeRequester").then(() => { }).catch(() => { }); }} />
                          </div>
                          <Separator />
                        </div>
                      )}
                      <div id="authorSection" className={" row"} >
                        <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +style.fieldCols + " col-lg-6 col-md-6 col-sm-12"}>
                          <TextField id="title" label="Title" defaultValue={this.state.title} value={this.state.title} onChange={this._onControlChange} styles={textFieldStyles} required={true} />
                        </div>
                        <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +style.fieldCols + " col-lg-6 col-md-6 col-sm-12"}>
                          <Dropdown
                            id="documentType"
                            required={true}
                            label="Document Type"
                            selectedKey={this.state.selectedDocumentType !== null ? this.state.selectedDocumentType : undefined}
                            onChange={this._onControlChange}
                            placeholder="Select Document Type"
                            options={this.state.dropdownDocumentTypeOptions}
                            styles={dropdownStyles}
                          />
                        </div>
                        <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
                          <TextField readOnly={this.state.isDisableEditing } id="description" label="Description" defaultValue={this.state.description} value={this.state.description} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                        </div>
                        <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +style.fieldCols + " col-lg-6 col-md-6 col-sm-12"}>
                          <Dropdown
                            id="department"
                            required={true}
                            label="Department"
                            selectedKey={this.state.selectedDepartment !== null ? this.state.selectedDepartment : undefined}
                            onChange={this._onControlChange}
                            placeholder=""
                            options={this.state.dropdownDepartmentOptions}
                            styles={dropdownStyles}
                          />
                        </div>
                        <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +style.fieldCols + " col-lg-6 col-md-6 col-sm-12 " + (this.state.isDisableIDCOSelection ? style.disabled : "")}>
                          <Dropdown
                            id="isIDCO"
                            required={true}
                            label="Is IDCO?"
                            selectedKey={this.state.is_idco}
                            onChange={this._onControlChange}
                            placeholder="Is IDCO"
                            options={[{ key: "Yes", text: "Yes" }, { key: "No", text: "No" }]}
                            styles={dropdownStyles}
                          />
                        </div>
                        {this.state.is_idco === "Yes" && (
                          <>
                            <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +style.fieldCols + " col-lg-6 col-md-6 col-sm-12 " + (this.state.isDisableIDCOSelection ? style.disabled : "")} >
                              <Dropdown
                                id="idco_val"
                                required={true}
                                label="IDCO"
                                selectedKey={this.state.selectedIDCO !== null ? this.state.selectedIDCO : undefined}
                                onChange={this._onControlChange}
                                placeholder="Select IDCO"
                                options={this.state.dropdownIDCOOptions}
                                styles={dropdownStyles}
                              />
                            </div>
                            <div className={style.fieldCols + " col-lg-6 col-md-6 col-sm-12"}>
                              <TextField id="idco_description" label="IDCO Description" value={this.state.selectedIDCOObj !== null ? this.state.selectedIDCOObj.Description : ""} readOnly styles={textFieldStyles} />
                            </div>
                            {/* <div className={style.fieldCols + " col-lg-6 col-md-6 col-sm-12"}>
                              <PeoplePicker
                                context={this.props.context as any}
                                titleText="IDCO Approver"
                                ensureUser={true}
                                personSelectionLimit={1}
                                onChange={this._onIDCOChange}
                                showtooltip={true}
                                required={true}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                defaultSelectedUsers={this.state.selectedIDCOObj !== null ? [this.state.selectedIDCOObj.ApproverEmail] : []}
                                resolveDelay={1000}
                                styles={peoplePickerStyles} />
                            </div> */}
                          </>
                        )}
                        {/* { (this.state.isOtherFunctionApprovalRequired === "Yes") && ( */}
                        <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +"col-12"}>
                          <label className='labelText' style={{ fontSize: '14px', fontWeight: 600, margin: '5px 0px' }}>Other Functions Approvers</label>
                          <div className='row'>
                            <div className={style.fieldCols + " col-lg-4 col-md-4 col-sm-12"}>
                              <PeoplePicker
                                peoplePickerWPclassName={style.OFPeoplePicker}
                                context={this.props.context as any}
                                titleText="Global Quality"
                                ensureUser={true}
                                personSelectionLimit={1}
                                onChange={this._onOFApproverChangeGQ}
                                showtooltip={true}
                                required={true}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                defaultSelectedUsers={this.state.selectedOFApproverGQ !== null ? [this.state.selectedOFApproverGQ.ApproverEmail] : []}
                                resolveDelay={1000}
                                styles={OFpeoplePickerStyles} />
                            </div>
                            <div className={style.fieldCols + " col-lg-4 col-md-4 col-sm-12"} style={{ paddingRight: '0px', paddingLeft: '0px' }}>
                              <PeoplePicker
                                peoplePickerWPclassName={style.OFPeoplePicker}
                                context={this.props.context as any}
                                titleText="Global Manufacturing Engineering"
                                ensureUser={true}
                                personSelectionLimit={1}
                                onChange={this._onOFApproverChangeGME}
                                showtooltip={true}
                                required={true}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                defaultSelectedUsers={this.state.selectedOFApproverGME !== null ? [this.state.selectedOFApproverGME.ApproverEmail] : []}
                                resolveDelay={1000}
                                styles={OFpeoplePickerStyles} />
                            </div>
                            <div className={style.fieldCols + " col-lg-4 col-md-4 col-sm-12"}>
                              <PeoplePicker
                                peoplePickerWPclassName={style.OFPeoplePicker}
                                context={this.props.context as any}
                                titleText="Global Purchasing"
                                ensureUser={true}
                                personSelectionLimit={1}
                                onChange={this._onOFApproverChangeGP}
                                showtooltip={true}
                                required={true}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                defaultSelectedUsers={this.state.selectedOFApproverGP !== null ? [this.state.selectedOFApproverGP.ApproverEmail] : []}
                                resolveDelay={1000}
                                styles={OFpeoplePickerStyles} />
                            </div>
                          </div>
                        </div>
                        {/* )} */}

                        <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +(this.state.isDisableSMESelection ? style.disabled : "") + " col-lg-6 col-md-6 col-sm-6 " + style.fieldCols}>
                          <Dropdown
                            id="isSMEEvaluationRequired"
                            required={true}
                            label="SME Evaluation Required?"
                            selectedKey={this.state.isSMEEvaluationRequired}
                            onChange={this._onControlChange}
                            placeholder="SME Evaluation Required"
                            options={[{ key: "Yes", text: "Yes" }, { key: "No", text: "No" }]}
                            styles={dropdownStyles}
                          />
                        </div>
                        {this.state.isSMEEvaluationRequired === "Yes" && (
                          <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
                            <PeoplePicker
                              context={this.props.context as any}
                              titleText="SME (Please refer GD000006876 for SMEs names list)"
                              ensureUser={true}
                              personSelectionLimit={100}
                              showtooltip={true}
                              required={true}
                              showHiddenInUI={false}
                              principalTypes={[PrincipalType.User]}
                              onChange={this._onSMEChange}
                              defaultSelectedUsers={this.state.selectedSME !== null ? this.state.defaultSME : []}
                              resolveDelay={1000}
                              styles={peoplePickerStyles} />
                          </div>
                        )}
                        {((this.state.folderId !== null && this.state.isShowUploadOption) && (this.state.isAuthorLoggedIn || this.state.isAdminLoggedIn)) &&
                          (
                            <div className={(this.state.isDisableEditing ? style.disabled : "") +" " +style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
                              <Label required>Upload Document</Label>
                              <input className={style.inputFile} type="file" ref={this.importFileUploadRef} onChange={this._onAttachmentFileChange} multiple={false} />
                              <CommandBarButton className={style.button} primary={true} text="Upload!" onClick={this._onAttachmentFileUpload} styles={buttonStyles} />
                            </div>
                          )
                        }

                      </div>
                    </div>
                  </div>
                  <Separator />
                  <>
                    {(this.state.isUnderSMEEvaluation && this.state.isSMELoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <>
                            <div id="smeSection" className='row'>
                              <div className={style.fieldCols + " col-12"}>
                                <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                                <div className={style.btnRow}>
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnSendBack} primary={true} text="Send back to requester" onClick={() => { this._onApproverAction("SME Evaluation", "Sent back to requester").then(() => { }).catch(() => { }); }} />
                                  {/* <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnReject} primary={true} text="Reject" onClick={() => { this._onApproverAction("SME Evaluation", "Rejected").then(() => { }).catch(() => { }); }} /> */}
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Complete" onClick={() => { this._onApproverAction("SME Evaluation", "Approved").then(() => { }).catch(() => { }); }} />
                                </div>
                              </div>
                            </div>
                            <Separator />
                          </>
                        )}
                      </>
                    )}
                    {(this.state.isInApprovalSME && this.state.isSMELoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <>
                            <div id="smeSection" className='row'>
                              <div className={style.fieldCols + " col-12"}>
                                <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                                <div className={style.btnRow}>
                                  {/* <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnSendBack} primary={true} text="Send back to requester" onClick={() => { this._onApproverAction("SME", "Sent back to requester").then(() => { }).catch(() => { }); }} /> */}
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnReject} primary={true} text="Reject" onClick={() => { this._onApproverAction("SME", "Rejected").then(() => { }).catch(() => { }); }} />
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Approve" onClick={() => { this._onApproverAction("SME", "Approved").then(() => { }).catch(() => { }); }} />
                                </div>
                              </div>
                            </div>
                            <Separator />
                          </>
                        )}
                      </>
                    )}
                    {(this.state.isInApprovalOtherFunction && this.state.isOtherFunctionApproverLoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <>
                            <div id="smeSection" className='row'>
                              <div className={style.fieldCols + " col-12"}>
                                <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                                <div className={style.btnRow}>
                                  {/* <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnSendBack} primary={true} text="Send back to requester" onClick={() => { this._onApproverAction("Other Functions", "Sent back to requester").then(() => { }).catch(() => { }); }} /> */}
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnReject} primary={true} text="Reject" onClick={() => { this._onApproverAction("Other Functions", "Rejected").then(() => { }).catch(() => { }); }} />
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Approve" onClick={() => { this._onApproverAction("Other Functions", "Approved").then(() => { }).catch(() => { }); }} />
                                </div>
                              </div>
                            </div>
                            <Separator />
                          </>
                        )}
                      </>
                    )}
                    {(this.state.isInApprovalCX && this.state.isCXApproverLoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <>
                            <div id="smeSection" className='row'>
                              <div className={style.fieldCols + " col-12"}>
                                <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                                <div className={style.btnRow}>
                                  {/* <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnSendBack} primary={true} text="Send back to requester" onClick={() => { this._onApproverAction("CX", "Sent back to requester").then(() => { }).catch(() => { }); }} /> */}
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnReject} primary={true} text="Reject" onClick={() => { this._onApproverAction("CX", "Rejected").then(() => { }).catch(() => { }); }} />
                                  <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Approve" onClick={() => { this._onApproverAction("CX", "Approved").then(() => { }).catch(() => { }); }} />
                                </div>
                              </div>
                            </div>
                            <Separator />
                          </>
                        )}
                      </>
                    )}
                  </>
                 
                    <div className={"row " + style.sectionButtons}>
                      <div className={style.fieldCols + " col-lg-12 col-md-12 col-sm-12"}>
                        {((this.state.isAuthorLoggedIn || this.state.isAdminLoggedIn) && this.state.isWorkflowInProgress !== "Yes") &&
                          (
                            <>
                              {this.state.isDraftStage && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Save" onClick={() => { this._onsubmit("EditRequest-Draft").then(() => { }).catch(() => { }); }} />
                              )}
                              {((this.state.isDraftStage && this.state.isSMEEvaluationRequired === "Yes") || this.state.isSMEEvaluationRejected) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Send for SME Evaluation" onClick={() => { this._onsubmit("Send For SME Evaluation").then(() => { }).catch(() => { }); }} />
                              )}
                              {((this.state.isUnderSMEEvaluation && this.state.isShowSMEEvaluationTab) || ((this.state.isSMEEvaluationDone || this.state.isSMERejected || this.state.isCXRejected || this.state.isOtherFunctionsRejected) && this.state.isSMEEvaluationRequired === "Yes")) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Send for SME Approval" onClick={() => { this._onsubmit("Send For SME Approval").then(() => { }).catch(() => { }); }} />
                              )}
                              {((this.state.isSMEApproved || this.state.isOtherFunctionsRejected) || (this.state.isDraftStage && this.state.isSMEEvaluationRequired === "No") || (this.state.isCXRejected)) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Send for Other Functions Approval" onClick={() => { this._onsubmit("Send for Other Functions Approval").then(() => { }).catch(() => { }); }} />
                              )}
                              {((this.state.isOtherFunctionsApproved || this.state.isCXRejected) || (this.state.isDraftStage && this.state.isSMEEvaluationRequired === "No" && this.state.isAdminLoggedIn)) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Send for CX Approval" onClick={() => { this._onsubmit("Send for CX Approval").then(() => { }).catch(() => { }); }} />
                              )}
                              {this.state.isCXApproved && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Send for Closure / Approved" onClick={() => { this._onsubmit("Send for Closure / Approved").then(() => { }).catch(() => { }); }} />
                              )}
                              {(this.state.isCXRejected || this.state.isOtherFunctionsRejected || this.state.isSMERejected) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Send for Closure / Rejected" onClick={() => { this._onsubmit("Send for Closure / Rejected").then(() => { }).catch(() => { }); }} />
                              )}
                              {(this.state.isSentForClosureApproved && this.state.isAdminLoggedIn) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Close / Approved" onClick={() => { this._onsubmit("Close / Approved").then(() => { }).catch(() => { }); }} />
                              )}
                              {(this.state.isSentForClosureRejected && this.state.isAdminLoggedIn) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Close / Rejected" onClick={() => { this._onsubmit("Close / Rejected").then(() => { }).catch(() => { }); }} />
                              )}
                              {(this.state.isSentForClosureExpired && this.state.isAdminLoggedIn) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Close / Expired" onClick={() => { this._onsubmit("Close / Expired").then(() => { }).catch(() => { }); }} />
                              )}
                              {(this.state.isAdminLoggedIn && !this.state.isClosed) && (
                                <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Close / Cancelled" onClick={() => { this._onsubmit("Close / Cancelled").then(() => { }).catch(() => { }); }} />
                              )}
                            </>
                          )}
                        <CommandBarButton style={{ display: 'inline-block' }} styles={buttonStyles} className={style.button} primary={true} text="Exit" onClick={() => { window.location.href = this.props.context.pageContext.web.absoluteUrl }} />
                      </div>
                    </div>
                 
                  {/* <Separator /> */}
                </PivotItem>
                {(this.state.isShowSMEEvaluationTab && this.state.isSMEEvaluationRequired === "Yes") && (
                  <PivotItem headerText="SME Evaluation">
                    {/* {(this.state.isUnderSMEEvaluation && this.state.isSMELoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <div id="smeSection" className='row'>
                            <div className={style.fieldCols + " col-12"}>
                              <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                              <div className={style.btnRow}>
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnSendBack} primary={true} text="Send back to requester" onClick={() => { this._onApproverAction("SME Evaluation", "Sent back to requester").then(() => { }).catch(() => { }); }} />
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Complete" onClick={() => { this._onApproverAction("SME Evaluation", "Approved").then(() => { }).catch(() => { }); }} />
                              </div>
                            </div>

                            
                          </div>
                        )}
                      </>
                    )} */}
                    {/* {(this.state.isSMELoggedIn || this.state.isAuthorLoggedIn || !this.state.isUnderSMEEvaluation) && ( */}
                    <>
                      <Label className={style.approvalHistoryLabel}>Approval History</Label>
                      <Separator />
                      <DetailsList
                        items={this.state.approvalHistorySMEEvaluation}
                        compact={false}
                        columns={this.state.approvalHistorySMEEvaluationColumns}
                        selectionMode={SelectionMode.none}
                        getKey={this._getKey}
                        setKey="none"
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        onRenderItemColumn={this.onRenderItemColumn}
                      />
                    </>
                    {/* )} */}
                  </PivotItem>
                )}
                {(this.state.isShowSMETab && this.state.isSMEEvaluationRequired === "Yes") && (
                  <PivotItem headerText="SME Approval">
                    {/* {(this.state.isInApprovalSME && this.state.isSMELoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <div id="smeSection" className='row'>
                            <div className={style.fieldCols + " col-12"}>
                              <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                              <div className={style.btnRow}>
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnReject} primary={true} text="Reject" onClick={() => { this._onApproverAction("SME", "Rejected").then(() => { }).catch(() => { }); }} />
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Approve" onClick={() => { this._onApproverAction("SME", "Approved").then(() => { }).catch(() => { }); }} />
                              </div>
                            </div>
                          </div>
                        )}
                      </>
                    )} */}
                    {/* {(this.state.isSMELoggedIn || this.state.isAuthorLoggedIn || !this.state.isInApprovalSME) && ( */}
                    <>
                      <Label className={style.approvalHistoryLabel}>Approval History</Label>
                      <Separator />
                      <DetailsList
                        items={this.state.approvalHistorySME}
                        compact={false}
                        columns={this.state.approvalHistorySMEColumns}
                        selectionMode={SelectionMode.none}
                        getKey={this._getKey}
                        setKey="none"
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        onRenderItemColumn={this.onRenderItemColumn}
                      />
                    </>
                    {/* )} */}
                  </PivotItem>
                )}
                {(this.state.isShowOtherFunctionsTab && this.state.isOtherFunctionApprovalRequired === "Yes") && (
                  <PivotItem headerText="Other Functions">
                    {/* {(this.state.isInApprovalOtherFunction && this.state.isOtherFunctionApproverLoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <div id="smeSection" className='row'>
                            <div className={style.fieldCols + " col-12"}>
                              <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                              <div className={style.btnRow}>
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnReject} primary={true} text="Reject" onClick={() => { this._onApproverAction("Other Functions", "Rejected").then(() => { }).catch(() => { }); }} />
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Approve" onClick={() => { this._onApproverAction("Other Functions", "Approved").then(() => { }).catch(() => { }); }} />
                              </div>
                            </div>                           
                          </div>
                        )}
                      </>
                    )} */}
                    {/* {(this.state.isOtherFunctionApproverLoggedIn || this.state.isAuthorLoggedIn || !this.state.isInApprovalOtherFunction) && ( */}
                    <>
                      <Label className={style.approvalHistoryLabel}>Approval History</Label>
                      <Separator />
                      <DetailsList
                        items={this.state.approvalHistoryOtherFunctions}
                        compact={false}
                        columns={this.state.approvalHistoryOtherFunctionsColumns}
                        selectionMode={SelectionMode.none}
                        getKey={this._getKey}
                        setKey="none"
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        onRenderItemColumn={this.onRenderItemColumn}
                      />
                    </>
                    {/* )} */}
                  </PivotItem>
                )}
                {this.state.isShowCXTab && (
                  <PivotItem headerText="CX">
                    {/* {(this.state.isInApprovalCX && this.state.isCXApproverLoggedIn) && (
                      <>
                        {!this.state.approvalDone && (
                          <div id="smeSection" className='row'>
                            <div className={style.fieldCols + " col-12"}>
                              <TextField id="reviewComments" label="Review Comments" defaultValue={this.state.reviewComments} value={this.state.reviewComments} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                              <div className={style.btnRow}>
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnReject} primary={true} text="Reject" onClick={() => { this._onApproverAction("CX", "Rejected").then(() => { }).catch(() => { }); }} />
                                <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button + " " + style.btnApprove} primary={true} text="Approve" onClick={() => { this._onApproverAction("CX", "Approved").then(() => { }).catch(() => { }); }} />
                              </div>
                            </div>
                            
                          </div>
                        )}
                      </>
                    )} */}
                    {/* {(this.state.isCXApproverLoggedIn || this.state.isAuthorLoggedIn || !this.state.isInApprovalCX) && ( */}
                    <>
                      <Label className={style.approvalHistoryLabel}>Approval History</Label>
                      <Separator />
                      <DetailsList
                        items={this.state.approvalHistoryCX}
                        compact={false}
                        columns={this.state.approvalHistoryCXColumns}
                        selectionMode={SelectionMode.none}
                        getKey={this._getKey}
                        setKey="none"
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        onRenderItemColumn={this.onRenderItemColumn} />
                    </>
                    {/* )} */}
                  </PivotItem>
                )}
                <PivotItem headerText="Discussions">
                  <div>
                    <TextField id="discussionComment" label="Post Comment" defaultValue={this.state.discussionComment} value={this.state.discussionComment} multiline rows={1} onChange={this._onControlChange} styles={textFieldStyles} />
                    <div className={style.btnRow}>
                      <CommandBarButton styles={buttonStyles} style={{ display: 'inline-block' }} className={style.button} primary={true} text="Post" onClick={() => { this._onPostDiscussion().then(() => { }).catch(() => { }); }} />
                    </div>
                  </div>
                  <div className={style.discussionItems}>
                    {this.state.discussions.map(discussionItem => {
                      return (
                        <div className={style.discussionItemDiv}>
                          <div className={style.discussionAuthorDetailsDiv}>
                            <img className={style.discussionAuthorImg} src={this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?size=M&accountname=" + discussionItem.Author.EMail}></img>
                            <div className={style.discussionAuthor}>
                              <p className={style.authorName}>{discussionItem.Author.Title}</p>
                              <p className={style.discussionCreated}>{moment(discussionItem.Created).format('DD MMM YYYY hh:mm A')}</p>
                            </div>
                          </div>
                          <div className={style.discussionComment}>{discussionItem.Comment}</div>
                          <Separator />
                        </div>

                      )
                    })}
                  </div>
                </PivotItem>
              </Pivot>

            </div>
          </div>
        </div>
      </Fabric >
    )
  }
  public rednderLoader = (): JSX.Element => {
    //Shows loader
    return (
      <div>
        {/* <Label>Spinner with label positioned below</Label> */}
        <Spinner label="Please wait.. Form is loading..." />
      </div>
    )
  }
  public renderNotAuthorisedForm = (): JSX.Element => {
    //Render not authoised form
    return (
      <Fabric className={style.boxShadow}>
        <div className={style.gridContainer}>
          <div className={style.topbanner + " row"}>
            <div className="col-sm-12" style={{ textAlign: "right" }}><span className='p-2'>Welcome, <b>{this.state.currentUser !== null ? this.state.currentUser.Title : <b></b>}</b></span></div>
          </div>
          <div className='row'>
            <div className="col-sm-12">
              <p>You are not authorised to view this initiative</p>
            </div>
          </div>
        </div>
      </Fabric >
    )
  }
}
