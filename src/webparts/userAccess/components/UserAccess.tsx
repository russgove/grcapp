import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import styles from './UserAccess.module.scss';
import { IUserAccessProps } from './IUserAccessProps';
import { IUserAccessState } from './IUserAccessState';
import { escape } from '@microsoft/sp-lodash-subset';
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
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
    super(props);
    console.log("in Construrctor");
    initializeIcons();
    this.selection.getKey = (item => { return item["ID"]; });
    this.state = {
      primaryApprover: null,
      userAccessItems: [],
      showTcodePopup: false,
      showApprovalPopup: false
    };
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

 
 
  public showPopup(item: UserAccessItem) {
    this.fetchRoleToTransaction(item.Role)
      .then((roleToTransactions) => {
        this.setState((current) => ({ ...current, roleToTransaction: roleToTransactions, showTcodePopup: true }));
      })
     .catch((err) => {
       console.error(err.data.responseBody["odata.error"].message.value);
       alert(err.data.responseBody["odata.error"].message.value);
       debugger;
    });
  }
/**
 * 
 * This method gets called by the popup window to update all the selected items
 */
  @autobind
  public updateSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let  tempArray:UserAccessItem[] = map(this.state.userAccessItems, (uaItem) => {
      if (this.selection.isKeySelected(uaItem.ID.toString()) === this.state.changeSelected) {
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
  @autobind
  public getApi(controller: string, query: string): Promise<any> {
    let url = this.props.webApiUrl + "/api/" + controller + "?" + query;
    let httpClientOptions: IHttpClientOptions = {
      credentials: "include",
    };
    return this.props.httpClient.get(url,    HttpClient.configurations.v1,httpClientOptions)
      .then((response: HttpClientResponse) => {
        return response.json();
      });
  }
  @autobind
  public putApi(controller: string, entity: any): Promise<any> {
    let url = this.props.webApiUrl + "/api/" + controller + "/" + entity["ID"];
    let requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    let httpClientOptions: IHttpClientOptions = {
      credentials: "include",
      headers: requestHeaders,
      method: "PUT",
      body: JSON.stringify(entity)
    };
    return this.props.httpClient.fetch(url, HttpClient.configurations.v1, httpClientOptions);
  }
  @autobind
  public updateUserAccessItems(items: UserAccessItem[]): Promise<any> {
    let promises: Array<Promise<any>> = [];
    for (let item of items) {
      promises.push(this.putApi(this.props.userAccessController, item));
    }
    return Promise.all(promises);
  }
 
  @autobind
  public setComplete(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let updatedApprover = this.state.primaryApprover;
    updatedApprover.Completed = "Yes";
    this.putApi(this.props.primaryApproverController, updatedApprover)
      .then(() => {
        this.setState((current) => ({ ...current, primaryApprover: updatedApprover }));
      })
      .catch((err) => {
        console.log(err);
        alert("An error occurred saving the primary approver record");
      });
  }
  @autobind
  public unsetComplete(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let updatedApprover = this.state.primaryApprover;
    updatedApprover.Completed = "";
    this.putApi(this.props.primaryApproverController, updatedApprover)
      .then(() => {
        this.setState((current) => ({ ...current, primaryApprover: updatedApprover }));
      })

      .catch((err) => {
        console.log(err);
        alert("An error occurred saving the primary approver record");
      });

  }
  @autobind
  public save(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    this.updateUserAccessItems(this.state.userAccessItems).then(() => {
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
  @autobind
  public addApprover(approver: any): Promise<HttpClientResponse> {
    
    let requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    requestHeaders.append('Cache-Control', 'no-cache');

    let httpClientOptions: IHttpClientOptions = {
      credentials: "include",
      body: JSON.stringify(approver)
    };

    //let url=`${this.props.webApiUrl}/api/${this.props.primaryApproverController}?$filter=tolower(ApproverEmail) eq '${approverEmail.toLowerCase()}'`;
    let url = `${this.props.webApiUrl}/api/${this.props.primaryApproverController}`;
    console.log(url);
    return this.props.httpClient.post(url,
      HttpClient.configurations.v1,
      httpClientOptions);
  }
  
  @autobind
  public fetchPrimaryApprover(): Promise<any> {
       let query = "$filter=tolower(ApproverEmail) eq '" + this.props.user.email.toLowerCase() + "'";
    return this.getApi(this.props.primaryApproverController, query)
      .then((appr) => {
        debugger;
        this.setState((current) => ({ ...current, primaryApprover: appr[0] }));
      }).catch(e => {
        console.log(e);
        alert("There was an error fetching Primary approvers");
      });

  }
  @autobind
  public fetchUserAccess(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let query = "$filter=tolower(ApproverEmail) eq '" + this.props.user.email.toLowerCase() + "'";
     this.getApi(this.props.userAccessController, query)
      .then((response: any) => {
        this.setState((current) => ({ ...current, userAccessItems: response }));
      })
      .catch(err => {
        console.log(err);
        alert("There was an error fetching User Access Items");
      });
  }
  @autobind
  public fetchRoleToTransaction(role: string) {
    let query = "$filter=Composite_role eq '" + role + "'";
    return this.getApi(this.props.roleToTcodeController, query);
  }

  /**
   * this function gets called after the iframe has connected to the webapi.
   * After this we can make calls to the web api passing the credentials
   */
  public frameLoaded() {

    this.fetchUserAccess();
    this.fetchPrimaryApprover();
  }

  public RenderApproval(item?: UserAccessItem, index?: number, column?: IColumn): JSX.Element {
    let options = [
      { key: "1", text: "Yes" },
      { key: "2", text: "No" }
    ];
    if (this.state.primaryApprover.Completed === "Yes") {
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
              return x.ID === item.ID;
            });
            rtc.Approval = option.key as string;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, roleToTCodeReview: tempRoleToTCodeReview, changesHaveBeenMade: true }));
          }}
        />
      );
    }
  }
  public RenderComments(item?: UserAccessItem, index?: number, column?: IColumn): JSX.Element {
    if (this.state.primaryApprover.Completed === "Yes") {
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
              return x.ID === item.ID;
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
  public render(): React.ReactElement<IUserAccessProps> {

    debugger;
    let itemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "Update Selected",
        name: "Update Selected",
        icon: "TriggerApproval",
        disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes",
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
        disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes",
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
        key: "Done", 
        name: "Complete",
         icon: "Completed",
          onClick: this.setComplete,
        disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes" ||
        (filter(this.state.userAccessItems, (rr) => { return rr.Approval === "3"; }).length > 0) // "3" is the initial state after larry uploads the access db
      }      ,
      { 
         key: "UnDone", name: "UnComplete", icon: "Completed", onClick: this.unsetComplete
      }

    ];
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {

        key: "Save", name: "Save", icon: "Save", onClick: this.save,
        disabled: !(this.state.primaryApprover) || !(filter(this.state.userAccessItems, (rr) => { return rr.hasBeenUpdated; }).length > 0)
        || this.state.primaryApprover.Completed === "Yes"

      }
    ];


    return (
      <div className={styles.userAccess}>
        <iframe src={this.props.webApiUrl} onLoad={this.frameLoaded.bind(this)} />

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
          key="ID"
          columns={[
            {
              key: "UserID", name: "User Id",
              fieldName: "User_ID", minWidth: 90, maxWidth: 90,
              isResizable: true,
            },
            {
              key: "userName", name: "User Name",
              fieldName: "User_Full_Name", minWidth: 100, maxWidth: 100,
              isResizable: true,
            },
            {
              key: "title", name: "Role Name",
              fieldName: "Role_name", minWidth: 300, maxWidth: 300,
              isResizable: true,

            },
            {
              key: "info", name: "Transactions",
              isResizable: true,
              fieldName: "Role_name", minWidth: 60, maxWidth: 60,
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
                fieldName: "Transaction_Text", minWidth: 150, maxWidth: 150,

              },

            ]}
          />
        </Panel>


      </div>
    );
  }
}
