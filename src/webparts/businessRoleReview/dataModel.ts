import { Item } from "sp-pnp-js";
// main list for grid
export  class BusinessRoleReviewItem extends Item {

    public Id: number;
    public Role_x0020_Name: string;
    public Composite_x0020_Role: string;
    public Approver:string;
    public Approver_x0020_Name:string;
    public ApproverEmail:string;
public Alt_x0020_Apprv:string; // this is the T-id
public Alternate_x0020_Approver:string; // this is the name


    public PrimaryApproverId:Number;
    public PrimaryApprover: {
        Title: string;
    };
    public Approval:string;
    public Comments:string;
    public Reviewed_x0020_by:string;
    public hasBeenUpdated: boolean; //set in code to trigger update'

}
export  class PrimaryApproverItem extends Item {

    public Id: number;
    public Owner_x0020_ID: string;
    public ApproverEmail: string;
    public Completed: string;
    public PrimaryApproverId: string;
    public PrimaryApprover: {
        Title: string;
    };

}