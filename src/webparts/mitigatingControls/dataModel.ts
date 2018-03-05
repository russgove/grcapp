import { Item } from "sp-pnp-js";
// main list for grid
export  class MitigatingControlsItem extends Item {

    public Id: number;
    public Control_x0020_ID: string;
    public Description: string;
    public Risk_x0020_ID: string;
    public Risk_x0020_Description: string;
    public Owner_x0020_ID: string;
    public ApproverEmail:string;
    public Control_x0020_Owner_x0020_Name:string;
    public Control_x0020_Monitor_x0020_ID:string;
    public Control_x0020_Monitor_x0020_Name:string;
    public Effective:string;
    public Continues:string;
    public Right_x0020_Monitor_x003f_:string;
    public Monitor_x0020_SOD_x0020_Conflict:string;
    public Date_x0020_Reviewed:string;
    public Cteam_x0020_Approver:string;
    public CTeam_x0020_Date:string;
    public Remediation:string;
    public CTeam_x0020_Comments:string;
    public Compliance_x0020_Team_x0020_ID:string;
    public Compliance_x0020_Team_x0020_Memb:string;
    public PrimaryApproverId:Number;
    public PrimaryApprover: {
        Title: string;
    };
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