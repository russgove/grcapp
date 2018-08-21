import { MitigatingControlsItem, PrimaryApproverItem } from "../dataModel";
export interface IMitigatingControlsState {
  primaryApprover : PrimaryApproverItem;
  mitigatingControls: Array<MitigatingControlsItem>;
  //roleToTransaction?: Array<RoleToTransaction>;
  showPopup:boolean;// triggers the popup that lets a user enter values for changeSelected
  popupValueEffective?:string; // th evalue entereed in the effective choicegroup on the popup
  popupValueContinues?:string; // th evalue entereed in the Continutes choicegroup on the popup
  popupValueCorrectPerson?:string; // th evalue entereed in the CorrectPerson choicegroup on the popup
  popupValueComments?:string; // th evalue entereed in the Comments textbox on the popup
  changeSelected?:boolean;// true changes selected items from pupup, false changes unselected

}
