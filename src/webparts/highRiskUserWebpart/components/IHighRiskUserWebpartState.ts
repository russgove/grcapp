import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import HighRisk from "../../../dataModel/HighRisk";
import RoleToTransaction from "../../../dataModel/RoleToTransaction"; 

export interface IHighRiskUserWebpartState {
  primaryApproverList: Array<PrimaryApproverList>;
  highRisk: Array<HighRisk>;
  roleToTransaction?: Array<RoleToTransaction>;
  showPopup:boolean;
}
