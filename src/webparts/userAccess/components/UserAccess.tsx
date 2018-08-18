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
import { Dialog, DialogFooter, DialogType } from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { find, map, clone, filter } from "lodash";
import { PrimaryApproverItem, UserAccessItem, RoleToTransaction } from "../datamodel";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';

export default class UserAccess extends React.Component<IUserAccessProps, IUserAccessState> {
  private selection: Selection = new Selection();
  public constructor(props: IUserAccessProps) {
    super();
    console.log("in Construrctor");
    initializeIcons();
    this.selection.getKey = (item => { return item["Id"]; });
    this.save = this.save.bind(this);
    this.setComplete = this.setComplete.bind(this);
    this.updateSelected = this.updateSelected.bind(this);
    this.fetchUserAccess = this.fetchUserAccess.bind(this);
    this.state = {
      primaryApproverList: props.primaryApproverList,
      userAccessItems: [],
      showTcodePopup: false,
      showApprovalPopup: false

    };
  }
  public componentDidMount() {
    debugger;
    this.props.fetchUserAccess()
    .then(userAccess => {
      debugger;
      this.setState((current) => ({ ...current, userAccessItems: userAccess }));
    })
    .catch((err)=>{
      console.log(err.data.responseBody["odata.error"].message.value);
      console.log(err);
      alert("error fething user Access Items")
      alert(err.data.responseBody["odata.error"].message.value);

      

    })
  }
  public componentDidUpdate(): void {
    // disable postback of buttons. see https://github.com/SharePoint/sp-dev-docs/issues/492
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
      { key: "1", text: "Yes" },
      { key: "2", text: "No" }

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
          value={item.Comments ? item.Comments : ""}
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
  public save( ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem) : void{
     this.props.save(this.state.userAccessItems).then(() => {

      var tempArray = map(this.state.userAccessItems, (rr) => {
        return { ...rr, hasBeenUpdated: false };
      });
      this.setState((current) => ({ ...current, userAccessItems: tempArray }));
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
  // public changeSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
  //   debugger;
  //   var tempArray = map(this.state.userAccessItems, (uaItem) => {
  //     if (this.selection.isKeySelected(uaItem.Id.toString())) {
  //       debugger;
  //       return {
  //         ...uaItem,
  //         Approval: this.state.popupValueApproval ? this.state.popupValueApproval : uaItem.Approval,
  //         Comments: this.state.popupValueComments ? this.state.popupValueComments : uaItem.Comments,
  //         hasBeenUpdated: true
  //       };
  //     }
  //     else {
  //       return {
  //         ...uaItem
  //       };
  //     }
  //   });
  //   this.setState((current) => ({ ...current, userAccessItems: tempArray }));
  // }
  // public changeUnSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
  //   debugger;
  //   var tempArray = map(this.state.userAccessItems, (uaItem) => {
  //     if (!this.selection.isKeySelected(uaItem.Id.toString())) {
  //       return {
  //         ...uaItem,
  //         Approval: this.state.popupValueApproval ? this.state.popupValueApproval : uaItem.Approval,
  //         Comments: this.state.popupValueComments ? this.state.popupValueComments : uaItem.Comments,
  //         hasBeenUpdated: true
  //       };
  //     }
  //     else {
  //       return {
  //         ...uaItem
  //       };
  //     }
  //   });
  //   this.setState((current) => ({ ...current, userAccessItems: tempArray }));
  // }
  // public changeAll(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {

  //   var tempArray = map(this.state.highRisk, (rr) => {
  //     return { ...rr, GRCApproval: item.data, hasBeenUpdated: true };
  //   });
  //   this.setState((current) => ({ ...current, highRisk: tempArray }));

  // }
  public showPopup(item: UserAccessItem) {

    this.props.getRoleToTransaction(item.Role)
      .then((roleToTransactions) => {

        this.setState((current) => ({ ...current, roleToTransaction: roleToTransactions, showTcodePopup: true }));

      }).catch((err) => {
        console.error(err.data.responseBody["odata.error"].message.value);
        alert(err.data.responseBody["odata.error"].message.value);

        debugger;
      });
  }

  public fetchUserAccess(): Promise<any> {

    return this.props.fetchUserAccess().then((items) => {
      debugger;
      this.setState((current) => ({ ...current, userAccessItems: items }));
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }
  public updateSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
debugger;
    var tempArray = map(this.state.userAccessItems, (uaItem) => {
      if (this.selection.isKeySelected(uaItem.Id.toString()) === this.state.changeSelected) {
        return {
          ...uaItem,
          Approval: this.state.popupValueApproval ? this.state.popupValueApproval : uaItem.Approval,
          Comments: this.state.popupValueComments ? this.state.popupValueComments : uaItem.Comments,
          hasBeenUpdated: true,
        };
      }
      else {
        return {
          ...uaItem
        };
      }
    });

    this.setState((current) => ({
      ...current,
      userAccessItems: tempArray,
      popupValueApproval: null,
      popupValueComments: null,
      showApprovalPopup: false
    }));
  }


  public render(): React.ReactElement<IUserAccessProps> {


    let itemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "Update Selected",
        name: "Update Selected",
        icon: "TriggerApproval",
        disabled: this.props.primaryApproverList[0].Completed === "Yes",
        onClick: (e) => {
          if (this.selection.count > 0) {
            this.setState((current) => ({
              ...current,
              showApprovalPopup: true,
              changeSelected: true
            }));
          }
        },

      },
      {
        key: "Update Unselected",
        name: "Update Unselected",
        icon: "TriggerAuto",
        disabled: !(this.props.primaryApproverList) || this.props.primaryApproverList[0].Completed === "Yes",
        onClick: (e) => {

          if (!this.selection.count || this.selection.count < this.state.userAccessItems.length) {
            this.setState((current) => ({
              ...current,
              showApprovalPopup: true,
              changeSelected: false // change UNSELECTED Items
            }));
          }
        },

      },

      {
        key: "Undo", name: "Undo", icon: "Undo", onClick: this.fetchUserAccess,
        disabled: !(filter(this.state.userAccessItems, (rr) => {
          return rr
            .hasBeenUpdated;
        }).length > 0)
      },
      { // if the item has been comleted OR there are items with noo approvasl, diable
        key: "Done", name: "Complete", icon: "Completed", onClick: this.setComplete,
        disabled: this.props.primaryApproverList[0].Completed === "Yes" ||
          (filter(this.state.userAccessItems, (rr) => { return rr.Approval === "3"; }).length > 0) // "3" is the initial state after larry uploads the access db
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
        <Dialog isBlocking={true}
          hidden={!this.state.showApprovalPopup}
          onDismiss={(e) => { this.setState((current) => ({ ...current, showApprovalPopup: false })); }}
          dialogContentProps={{
            type: DialogType.close,
            title: this.state.changeSelected
              ? `Update ${this.selection.count} Selected Items`
              : this.selection.count
                ? `Update ${this.state.userAccessItems.length - this.selection.count} Unselected Items`
                : `Update ${this.state.userAccessItems.length} Unselected Items`,
            subText: 'All selected items will be updated with the following values'
          }} >
          <ChoiceGroup label="Approval Decision"
            options={[
              {
                key: '1',
                text: 'Yes'
              },
              {
                key: '2',
                text: 'No',
              },
            ]}
            selectedKey={this.state.popupValueApproval}

            onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => {

              this.setState((current) => ({ ...current, popupValueApproval: option.key }));
            }}
          />

          <TextField label="Comments" onChanged={(e) => {

            this.setState((current) => ({ ...current, popupValueComments: e }));
          }}

          />



          <DialogFooter>
            <PrimaryButton text='Save' onClick={this.updateSelected.bind(this)} />
            <DefaultButton text='Cancel' onClick={(e) => {
              this.setState((current) => ({ ...current, showApprovalPopup: false }));
            }} />
          </DialogFooter>
        </Dialog>
        <CommandBar
          isSearchBoxVisible={false}
          items={itemsNonFocusable}
          farItems={farItemsNonFocusable}

        />
        <DetailsList
          items={this.state.userAccessItems}
          selectionMode={SelectionMode.multiple}
          selection={this.selection}
          key="Id"
          columns={[
            {
              key: "UserID", name: "User Id",
              fieldName: "User_x0020_ID", minWidth: 90, maxWidth: 90,
              isResizable: true,
            },
            {
              key: "userName", name: "User Name",
              fieldName: "User_x0020_Full_x0020_Name", minWidth: 100, maxWidth: 100,
              isResizable: true,
            },
            {
              key: "title", name: "Role Name",
              fieldName: "Role_x0020_name", minWidth: 300, maxWidth: 300,
              isResizable: true,

            },
            {
              key: "info", name: "Transactions",
              isResizable: true,
              fieldName: "Role_x0020_name", minWidth: 60, maxWidth: 60,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return (
                  <IconButton iconProps={{ iconName: "Info" }} onClick={(e) => { this.showPopup(item); }} />
                );

              }
            },
            {
              key: "Approval", name: "Approval",
              isResizable: true,
              fieldName: "Approval", minWidth: 90, maxWidth: 90,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderApproval(item, index, column); },
            },
            {
              key: "Comments", name: "Comments",
              fieldName: "Comments", minWidth: 150, maxWidth: 150,
              isResizable: true,
              onRender: (item?: any, index?: number, column?: IColumn) => { return this.RenderComments(item, index, column); },
            },


          ]}
        />

        <Panel
          type={PanelType.custom | PanelType.smallFixedNear}
          customWidth='900px'
          isOpen={this.state.showTcodePopup}
          onDismiss={
            () => {
              this.setState((current) => ({ ...current, roleToTransaction: [], showTcodePopup: false }));
            }}
          isBlocking={true}
        >
          <DetailsList
            items={this.state.roleToTransaction}
            selectionMode={SelectionMode.none}
            columns={[

              {
                key: "Role", name: "Role",
                isResizable: true,
                fieldName: "Role", minWidth: 300, maxWidth: 300,
              },
              {
                key: "Comments", name: "TCode",
                isResizable: true,
                fieldName: "TCode", minWidth: 50, maxWidth: 50,

              },
              {
                key: "Remediation", name: "Transaction Text",
                isResizable: true,
                fieldName: "Transaction_x0020_Text", minWidth: 150, maxWidth: 150,

              },

            ]}
          />
        </Panel>


      </div>
    );
  }
}
