
// main list for grid
export class RoleReviewItem  {

    public ID: number;
    public Role_Name: string;
    public Approver: string;
    public ApproverEmail: string;
    public Approver_Name: string;
    public Alt_Apprv: string;
    public AlternateApproverEmail:string;
    public Alternate_Approver:string;
    public Approved_by:string;
 
  
    public Approval: string;
    // public GRCApprovedBy: string;
    public Date_Reviewed: string;
    public Comments: string;
    public Remediation: string;
    public hasBeenUpdated: boolean;

}

export class RoleToTransaction  {
    public Role: string;
    public Composite_role: string;
    public TCode: string;
    public Transaction_Text: string;

}
export class PrimaryApproverItem  {

    public ID: number;
    public Approver: string;
    public ApproverEmail: string;
    public Approver_Name: string;
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