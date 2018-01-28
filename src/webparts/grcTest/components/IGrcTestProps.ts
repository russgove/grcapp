import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import RoleReview from "../../../dataModel/RoleReview";
import RoleToTransaction from "../../../dataModel/RoleToTransaction";

export interface IGrcTestProps {
  primaryApproverList: Array<PrimaryApproverList>;
   save: (roleToTCodeReview:  Array<RoleReview>) => Promise<any>; 
  getRoleToTransaction: (role:  string) => Promise<Array<RoleToTransaction>>; 
  fetchRoleReviews:() =>  Promise<Array<RoleReview>>; 
  setComplete: ( PrimaryApproverList:PrimaryApproverList) => Promise<any>; 
  
}
