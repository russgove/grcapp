import { Item } from "sp-pnp-js";
// tassks to be approved
export default class PrimaryApproverList extends Item {

    public Approver: string;
    public Approver_x0020_Name: string;
    public Completed: string;
    public AssignedToId: number;
    public AssignedTo: {
        Title: string;

    };
}