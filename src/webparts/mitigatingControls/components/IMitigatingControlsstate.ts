import { MitigatingControlsItem, PrimaryApproverItem } from "../dataModel";
export interface IMitigatingControlsState {
  primaryApprover : Array<PrimaryApproverItem>;
  mitigatingControls: Array<MitigatingControlsItem>;
  //roleToTransaction?: Array<RoleToTransaction>;
  showPopup:boolean;
}
