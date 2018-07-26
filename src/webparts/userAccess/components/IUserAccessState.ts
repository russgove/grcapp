import {PrimaryApproverItem,UserAccessItem,RoleToTransaction} from "../datamodel";

export interface IUserAccessState {
  primaryApprover: PrimaryApproverItem;
  userAccessItems: Array<UserAccessItem>;
  roleToTransaction?: Array<RoleToTransaction>;
  showTcodePopup:boolean;
  showApprovalPopup:boolean;
  popupValueApproval?:string; // the value entereed in the effective choicegroup on the popup
  popupValueComments?:string; // the value entereed in the Comments textbox on the popup
  changeSelected?:boolean;// true changes selected items from pupup, false changes unselected
}
