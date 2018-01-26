import { Item } from "sp-pnp-js";
// details to disp;lay when user clicks help
export default class RoleToTransaction extends Item {

    public Role_x0020_Name:string;
    public Composite_x0020_role:string;
    public Role:string;
    public TCode:string;
    public Transaction_x0020_Text:string;
    public Approver:string;
    public Approver_x0020_Name:string;
    public Alt_x0020_Apprv:string;
    public Alternate_x0020_Approver:string;
    public Approval:string;
    public Approved_x0020_By:string;
    public Date_x0020_Reviewed:string;
    public Comments:string;
    public Remediation:string;
}