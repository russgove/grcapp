import {PrimaryApproverItem,HighRiskFunction,RoleToTransaction} from "../datamodel";

export interface IHighRiskFunctuionsState {
  primaryApprover: PrimaryApproverItem;
  HighRiskFunctionItems: Array<HighRiskFunction>;
  roleToTransaction?: Array<RoleToTransaction>;
  showTcodePopup:boolean;
  showApprovalPopup:boolean;
  popupValueApproval?:string; // the value entereed in the effective choicegroup on the popup
  popupValueComments?:string; // the value entereed in the Comments textbox on the popup
  changeSelected?:boolean;// true changes selected items from pupup, false changes unselected
  showOverlay?:boolean;
  overlayMessage?:string;
}
