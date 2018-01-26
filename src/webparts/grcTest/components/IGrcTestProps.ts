import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import RoleReview from "../../../dataModel/RoleReview";
import RoleToTransaction from "../../../dataModel/RoleToTransaction";

export interface IGrcTestProps {
  primaryApproverList: Array<PrimaryApproverList>;
  roleReview: Array<RoleReview>;
  save: (roleToTCodeReview:  Array<RoleReview>) => Promise<any>; 
  getRoleToTransaction: (role:  string) => Promise<Array<RoleToTransaction>>; 
  setComplete: ( PrimaryApproverList:PrimaryApproverList) => Promise<any>; 
  
}
