import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import RoleToTCodeReview from "../../../dataModel/RoleToTCodeReview";

export interface IGrcTestState {
  primaryApproverList: Array<PrimaryApproverList>;
  roleToTCodeReview: Array<RoleToTCodeReview>;
  changesHaveBeenMade:boolean;
}
