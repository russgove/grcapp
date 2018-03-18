import {HighRiskItem, PrimaryApproverItem,RoleToTransaction } from "../dataModel";
export interface IHighRiskProps {
  primaryApprover: Array<PrimaryApproverItem>;
  save: (mitigatingControls: Array<HighRiskItem>) => Promise<any>;
  fetchHighRisks: () => Promise<Array<HighRiskItem>>;
  getRoleToTransaction: (role: string) => Promise<Array<RoleToTransaction>>;
  setComplete: (PrimaryApprover: PrimaryApproverItem) => Promise<any>;
  domElement: any; // needed to disable button postback after render on classic pages

}
