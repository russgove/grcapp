import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import RoleReview from "../../../dataModel/RoleReview";
import RoleToTransaction from "../../../dataModel/RoleToTransaction";

export interface IGrcTestState {
  primaryApproverList: Array<PrimaryApproverList>;
  roleReview: Array<RoleReview>;
  roleToTransaction?: Array<RoleToTransaction>;

  showPopup:boolean;

}
