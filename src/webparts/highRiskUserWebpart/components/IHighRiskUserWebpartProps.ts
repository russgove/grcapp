import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import HighRisk from "../../../dataModel/HighRisk";
import RoleToTransaction from "../../../dataModel/RoleToTransaction";

export interface IHighRiskUserWebpartProps {
  primaryApproverList: Array<PrimaryApproverList>;
  save: (highRisks: Array<HighRisk>) => Promise<any>;
  getRoleToTransaction: (role: string) => Promise<Array<RoleToTransaction>>;
  fetchHighRisk: () => Promise<Array<HighRisk>>;
  setComplete: (PrimaryApproverList: PrimaryApproverList) => Promise<any>;
  domElement: any; // needed to disable button postback after render on classic pages
}
