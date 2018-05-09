import * as React from 'react';
import styles from './UserAccess.module.scss';
import { IUserAccessProps } from './IUserAccessProps';
import { IUserAccessState } from './IUserAccessState';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import {
  DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, Selection,
  ColumnActionsMode
} from "office-ui-fabric-react/lib/DetailsList";
import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import { Modal, IModalProps } from "office-ui-fabric-react/lib/Modal";
import { Panel, IPanelProps, PanelType } from "office-ui-fabric-react/lib/Panel";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";

import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton, IconButton } from "office-ui-fabric-react/lib/Button";
import { Dialog } from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { find, map, clone, filter } from "lodash";
import {PrimaryApproverItem,UserAccessItem,RoleToTransaction} from "../datamodel"

export default class UserAccess extends React.Component<IUserAccessProps, IUserAccessState> {
  private selection: Selection = new Selection();
  public constructor(props: IUserAccessProps) {
    super();
    console.log("in Construrctor");
    initializeIcons();
    this.selection.getKey = (item => { return item["Id"]; });
    this.save = this.save.bind(this);
    this.setComplete = this.setComplete.bind(this);
    this.changeUnSelected = this.changeUnSelected.bind(this);
    this.changeSelected = this.changeSelected.bind(this);
    this.fetchUserAccess = this.fetchUserAccess.bind(this);
    this.state = {
      primaryApproverList: props.primaryApproverList,
      userAccessItems: [],
      showPopup: false

    };
  }
  public componentDidMount() {
    this.props.fetchUserAccess().then(userAccess => {
      debugger;
      this.setState((current) => ({ ...current, userAccessItems: userAccess }));

    });
  }
  public componentDidUpdate(): void {
    // disable postback of buttons. see https://github.com/SharePoint/sp-dev-docs/issues/492
    debugger;
    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      const buttons: NodeListOf<HTMLButtonElement> = this.props.domElement.getElementsByTagName('button');
      for (let i: number = 0; i < buttons.length; i++) {
        if (buttons[i]) {
          // Disable the button onclick postback
          buttons[i].onclick = () => {
            return false;
          };
        }
      }
    }
  }
  public RenderApproval(item?: UserAccessItem, index?: number, column?: IColumn): JSX.Element {

    let options = [
      { key: "0", text: "yup" },
      { key: "1", text: "nope" },
      { key: "2", text: "no f'in way" }
    ];
    if (this.props.primaryApproverList[0].Completed === "Yes") {
      return (
        <div>
          {find(options, (o) => { return o.key === item.Approval; }).text}
        </div>
      );
    }
    else {
      return (
        <Dropdown options={options}
          selectedKey={item.Approval}
          onChanged={(option: IDropdownOption, idx?: number) => {
            let tempRoleToTCodeReview = this.state.userAccessItems;

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


  }
  public RenderComments(item?: UserAccessItem, index?: number, column?: IColumn): JSX.Element {
    if (this.props.primaryApproverList[0].Completed === "Yes") {
      return (
        <div>
          {item.Comments}
        </div>
      );
    }
    else {
      return (
        <TextField
          value={item.Comments}
          onChanged={(newValue) => {
            let tempRoleToTCodeReview = this.state.userAccessItems;
            let rtc = find(tempRoleToTCodeReview, (x) => {
              return x.Id === item.Id;
            });
            rtc.Comments = newValue;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, roleToTCodeReview: tempRoleToTCodeReview, changesHaveBeenMade: true }));

          }}

        >

        </TextField>
      );
    }


  }
  public save(): Promise<any> {
    return this.props.save(this.state.userAccessItems).then(() => {
      debugger;
      var tempArray = map(this.state.userAccessItems, (rr) => {
        return { ...rr, hasBeenUpdated: false };
      });
      this.setState((current) => ({ ...current, highRisk: tempArray }));
      alert("Saved");
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }
  public setComplete(): Promise<any> {
    return this.props.setComplete(this.props.primaryApproverList[0]).then(() => {

      alert("Completed");
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }
  public changeSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    var tempArray = map(this.state.userAccessItems, (rr) => {
      if (this.selection.isKeySelected(rr.Id.toString())) {
        return {
          ...rr,
          Approval: item.data, hasBeenUpdated: true
        };
      }
      else {
        return {
          ...rr
        };
      }
    });
    this.setState((current) => ({ ...current, highRisk: tempArray }));
  }
  public changeUnSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    var tempArray = map(this.state.userAccessItems, (rr) => {
      if (!this.selection.isKeySelected(rr.Id.toString())) {
        return {
          ...rr,
          Approval: item.data, hasBeenUpdated: true
        };
      }
      else {
        return {
          ...rr
        };
      }
    });
    this.setState((current) => ({ ...current, highRisk: tempArray }));
  }
  // public changeAll(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {

  //   var tempArray = map(this.state.highRisk, (rr) => {
  //     return { ...rr, GRCApproval: item.data, hasBeenUpdated: true };
  //   });
  //   this.setState((current) => ({ ...current, highRisk: tempArray }));

  // }
  public showPopup(item: UserAccessItem) {

    this.props.getRoleToTransaction(item.Role)
      .then((roleToTransactions) => {

        this.setState((current) => ({ ...current, roleToTransaction: roleToTransactions, showPopup: true }));

      }).catch((err) => {
        console.error(err.data.responseBody["odata.error"].message.value);
        alert(err.data.responseBody["odata.error"].message.value);

        debugger;
      });
  }

  public fetchUserAccess(): Promise<any> {
    debugger;
    return this.props.fetchUserAccess().then((highrisks) => {
      debugger;
      this.setState((current) => ({ ...current, highRisk: highrisks }));
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }


  public render(): React.ReactElement<IUserAccessProps> {
   
    debugger;
    let itemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "Change Selected",
        name: "Change Selected",
        icon: "TriggerApproval",
        disabled: this.props.primaryApproverList[0].Completed === "Yes",
        subMenuProps: {
          items: [
            {
              key: 'yup',
              name: 'Yup',
              data: "0",
              onClick: this.changeSelected,
              disabled: this.props.primaryApproverList[0].Completed === "Yes"
            },
            {
              key: 'Nope',
              name: 'Nope',
              data: "1",
              onClick: this.changeSelected,
              disabled: this.props.primaryApproverList[0].Completed === "Yes"
            },
            {
              key: "no f'in way",
              name: "no f'in way",
              data: "2",
              onClick: this.changeSelected,
              disabled: this.props.primaryApproverList[0].Completed === "Yes"

            }

          ]
        }
      },
      {
        key: "Change Unselected",
        name: "Change Unselected",
        icon: "TriggerAuto",
        disabled: this.props.primaryApproverList[0].Completed === "Yes"  ,
        
        subMenuProps: {
          items: [
            {
              key: 'yup',
              name: 'Yup',
              data: "0",
              onClick: this.changeUnSelected,
              disabled: this.props.primaryApproverList[0].Completed === "Yes" 
            
            },
            {
              key: 'Nope',
              name: 'Nope',
              data: "1",
              onClick: this.changeUnSelected
            },
            {
              key: "no f'in way",
              name: "no f'in way",
              data: "2",
              onClick: this.changeUnSelected

            }

          ]
        }
      },
   
      {
        key: "Undo", name: "Undo", icon: "Undo", onClick: this.fetchUserAccess,
        disabled: !(filter(this.state.userAccessItems, (rr) => { return rr
          .hasBeenUpdated; }).length > 0)
      },
      { // if the item has been comleted OR there are items with noo approvasl, diable
        key: "Done", name: "Complete", icon: "Completed", onClick: this.setComplete,
        disabled: this.props.primaryApproverList[0].Completed === "Yes" ||
        (filter(this.state.userAccessItems, (rr) => { return rr.Approval === null; }).length > 0)
      }

    ];
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {

        key: "Save", name: "Save", icon: "Save", onClick: this.save,
        disabled: !(filter(this.state.userAccessItems, (rr) => { return rr.hasBeenUpdated; }).length > 0)
        || this.props.primaryApproverList[0].Completed === "Yes"

      }
    ];


    return (
      <div className={styles.userAccess}>
        <CommandBar
          isSearchBoxVisible={false}
          items={itemsNonFocusable}
          farItems={farItemsNonFocusable}

        />
        <IconButton iconProps={{ iconName: "Info" }} onClick={(e) => { this.showPopup(null); }} />
        <DetailsList
          items={this.state.userAccessItems}
          selectionMode={SelectionMode.multiple}
          selection={this.selection}
          key="Id"
          columns={[
               {
              key: "UserID", name: "User Id",
              fieldName: "User_x0020_ID", minWidth: 40, maxWidth: 40,
            },
            {
              key: "userName", name: "User Name",
              fieldName: "User_x0020_Full_x0020_Name", minWidth: 100, maxWidth: 100,
            },
            {
              key: "title", name: "Role Name",
              fieldName: "Role_x0020_name", minWidth: 300, maxWidth: 300,

            },
            {
              key: "info", name: "Info",
              fieldName: "Role_x0020_name", minWidth: 10, maxWidth: 10,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return (
                  <IconButton iconProps={{ iconName: "Info" }} onClick={(e) => { this.showPopup(item); }} />
                );
              }

            },
            {
              key: "Approval", name: "Approval",
              fieldName: "Approval", minWidth: 90, maxWidth: 90,
                onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderApproval(item, index, column); },
              },
            {
              key: "Comments", name: "Comments",
              fieldName: "Comments", minWidth: 150, maxWidth: 150,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderComments(item, index, column); },
            },
       

          ]}
        />

        <Panel
          type={PanelType.custom | PanelType.smallFixedNear}
          customWidth='600px'
          isOpen={this.state.showPopup}
          onDismiss={
            () => {
              this.setState((current) => ({ ...current, roleToTransaction: [], showPopup: false }));
            }}
          isBlocking={true}
        >
          <DetailsList
            items={this.state.roleToTransaction}
            selectionMode={SelectionMode.none}
            columns={[

              {
                key: "Role", name: "Role",
                fieldName: "Role", minWidth: 150, maxWidth: 150,
              },
              {
                key: "Comments", name: "TCode",
                fieldName: "TCode", minWidth: 50, maxWidth: 50,

              },
              {
                key: "Remediation", name: "Transaction Text",
                fieldName: "Transaction_x0020_Text", minWidth: 150, maxWidth: 150,

              },

            ]}
          />
        </Panel>


      </div>
    );
  }
}
