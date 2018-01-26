import { Item } from "sp-pnp-js";
// main list for grid
export default class RoleToTransaction extends Item {

    public Id:number;
    public Role_x0020_Name:string;    
    public Approver:string;
    public Approver_x0020_Name:string;
    public Alt_x0020_Apprv:string;
    public Alternate_x0020_Approver:string;
    public Approval:string;
    public Approved_x0020_By:string;
    public Date_x0020_Reviewed:string;
    public Comments:string;
    public Remediation:string;
    public hasBeenUpdated:boolean;
   
}