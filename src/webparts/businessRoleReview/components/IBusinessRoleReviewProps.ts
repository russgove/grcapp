import {BusinessRoleReviewItem, PrimaryApproverItem } from "../dataModel";
export interface IBusinessRoleReviewProps {
  primaryApprover: Array<PrimaryApproverItem>;
  save: (mitigatingControls: Array<BusinessRoleReviewItem>) => Promise<any>;
  fetchBusinessRoleReview: () => Promise<Array<BusinessRoleReviewItem>>;
  setComplete: (PrimaryApprover: PrimaryApproverItem) => Promise<any>;
  domElement: any; // needed to disable button postback after render on classic pages
  roleNameWidth:number;
  //compositeRoleWidth:number;
  approverWidth:number;
  altApproverWidth:number;
  approvalDecisionWidth:number;
  commentsWidth:number;

}
