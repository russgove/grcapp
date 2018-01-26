import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import RoleToTCodeReview from "../../../dataModel/RoleToTCodeReview";

export interface IGrcTestProps {
  primaryApproverList: Array<PrimaryApproverList>;
  roleToTCodeReview: Array<RoleToTCodeReview>;
  save: (roleToTCodeReview:  Array<RoleToTCodeReview>) => Promise<any>; 
}
