import { Item } from "@pnp/sp";
// main list for grid
export class RoleReviewItem extends Item {

    public ID: number;
    public Role_x0020_Name: string;
    public Approver: string;
    public ApproverEmail: string;
    public Approver_x0020_Name: string;
    public Alt_x0020_Apprv: string;
    public AlternateApproverEmail:string;
    public Alternate_x0020_Approver:string;
    public Approved_x0020_by:string;
 
  
    public Approval: string;
    // public GRCApprovedBy: string;
    public Date_x0020_Reviewed: string;
    public Comments: string;
    public Remediation: string;
    public hasBeenUpdated: boolean;

}

export class RoleToTransaction extends Item {
    public Role: string;
    public Composite_x0020_role: string;
    public TCode: string;
    public Transaction_x0020_Text: string;

}
export class PrimaryApproverItem extends Item {

    public ID: number;
    public Approver: string;
    public ApproverEmail: string;
    public Approver_x0020_Name: string;
    public Completed: string;


}
export class HelpLink {
    public Id: number; //id of the splistitem
    public Title: string; // library to store the items in
    public IconName: string; // the Reference # from the PBC list Mapped to Title in list
    public Url: {
        Description: string,
        Url: string
    };
    public Target: string; // period info is needed for
    public Specs: string; // date the user needs to upload the files by

}