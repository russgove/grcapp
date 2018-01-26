import * as React from 'react';
import styles from './GrcTest.module.scss';
import { IGrcTestProps } from './IGrcTestProps';
import { IGrcTestState } from './IGrcTestState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton } from "office-ui-fabric-react/lib/Button";

import PrimaryApproverList from "../../../dataModel/PrimaryApproverList";
import RoleToTCodeReview from "../../../dataModel/RoleToTCodeReview";
import { find } from "lodash";
export default class GrcTest extends React.Component<IGrcTestProps, IGrcTestState> {
  public constructor(props: IGrcTestProps) {
    super();
    console.log("in Construrctor");
    // this.CloseButton = this.CloseButton.bind(this);
    // this.CompleteButton = this.CompleteButton.bind(this);
    this.save = this.save.bind(this);
    this.state = {
      primaryApproverList: props.primaryApproverList,
      roleToTCodeReview: props.roleToTCodeReview,
      changesHaveBeenMade: false

    };
  }
  public RenderApproval(item?: RoleToTCodeReview, index?: number, column?: IColumn): JSX.Element {
  
    let options = [
      { key: "0", text: "yup" },
      { key: "1", text: "nope" },
      { key: "2", text: "no f'in way" }
    ];
    return (
      <Dropdown options={options}
        selectedKey={item.Approval}
        onChanged={(option: IDropdownOption, idx?: number) => {
          let tempRoleToTCodeReview = this.state.roleToTCodeReview;
        
          let rtc = find(tempRoleToTCodeReview, (x) => {
            return x.Id === item.Id;
          });
          rtc.Approval = option.key as string;
          rtc.hasBeenUpdated = true;
          this.setState((current) => ({ ...current, roleToTCodeReview: tempRoleToTCodeReview, changesHaveBeenMade: true }));

        }}

      >

      </Dropdown>
    );
  }
  public save(): Promise<any> {
    return this.props.save(this.state.roleToTCodeReview).then(() => {
      debugger;
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }
  public render(): React.ReactElement<IGrcTestProps> {
   
    return (
      <div className={styles.grcTest}>
        <PrimaryButton buttonType={ButtonType.primary} onClick={this.save} iconProps={{ iconName: "ms-Icon--Save" }}>
          <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
          Save
 </PrimaryButton>
        <DetailsList
          items={this.props.roleToTCodeReview}
          selectionMode={SelectionMode.none}
          columns={[
            {
              key: "title", name: "Role Name",
              fieldName: "Role_x0020_Name", minWidth: 400, maxWidth: 400

            },
            {
              key: "Approval", name: "Approval",
              fieldName: "Approval", minWidth: 90, maxWidth: 90,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderApproval(item, index, column); }

            },
            {
              key: "Comments", name: "Comments",
              fieldName: "Comments", minWidth: 150, maxWidth: 150,

            },
            {
              key: "Remediation", name: "Remediation",
              fieldName: "Remediation", minWidth: 150, maxWidth: 150,

            },

          ]}
        />
      </div>
    );
  }
}
