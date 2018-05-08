import { Item } from "@pnp/sp";
// tassks to be approved
export default class PrimaryApproverList extends Item {

        public Id:number;
    public GRCCompleted: string;
    public GRCApproverId: number;
    public GRCApprover: {
        Title: string;

    };
}