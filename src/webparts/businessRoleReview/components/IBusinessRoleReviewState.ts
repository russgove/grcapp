import { BusinessRoleReviewItem, PrimaryApproverItem } from "../dataModel";
export interface IBusinessRoleReviewState {
  primaryApprover : Array<PrimaryApproverItem>;
  businessRoleReview: Array<BusinessRoleReviewItem>;
  //roleToTransaction?: Array<RoleToTransaction>;
  showPopup:boolean;// triggers the popup that lets a user enter values for changeSelected
  popupValueApproval?:string; // th evalue entereed in the effective choicegroup on the popup
  popupValueComments?:string; // th evalue entereed in the Comments textbox on the popup
  changeSelected?:boolean;// true changes selected items from pupup, false changes unselected

}
