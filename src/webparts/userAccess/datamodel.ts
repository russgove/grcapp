import { Item } from "@pnp/sp";
// main list for grid
export class UserAccessItem extends Item {

    public Id: number;
    public User_x0020_ID: string;
    public User_x0020_Full_x0020_Name: string;
    public Role: string;
    public Role_x0020_name: string;
    public PrimaryApproverId: string;
    public PrimaryApprover: {
        Title: string;
    };
    // public AlternateApproverId: string;
    // public GRCAlternateApprover: {
    //     Title: string;
    // };
    public Approval: string;
   // public GRCApprovedBy: string;
    public Date_x0020_Reviewed: string;
    public Comments: string;
    public Remediation: string;
    public hasBeenUpdated: boolean;

}

export  class RoleToTransaction extends Item {
    public Role:string;
    public Composite_x0020_role:string;
    public TCode:string;
    public Transaction_x0020_Text:string;
    
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
export class HelpLink {
    public Id: number; //id of the splistitem
    public Title: string; // library to store the items in
    public IconName: string; // the Reference # from the PBC list Mapped to Title in list
    public Url: {
        Description:string,
        Url:string
    };
    public Target: string; // period info is needed for
    public Specs: string; // date the user needs to upload the files by
  
}