import styles from './RoleToTCode.module.scss';
import { IRoleToTCodeProps } from './IRoleToTCodeProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { IRoleToTCodeState } from './IRoleToTCodeState';
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
import { Modal, IModalProps } from "office-ui-fabric-react/lib/Modal";
import { Panel, IPanelProps, PanelType } from "office-ui-fabric-react/lib/Panel";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";
import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton, IconButton } from "office-ui-fabric-react/lib/Button";
import { Dialog, DialogFooter, DialogType } from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { find, map, clone, filter } from "lodash";
import { PrimaryApproverItem, RoleReviewItem, RoleToTransaction } from "../datamodel";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';


export default class RoleToTCode extends React.Component<IRoleToTCodeProps, IRoleToTCodeState> {
  private selection: Selection = new Selection();
  public constructor(props: IRoleToTCodeProps) {
    super();
    console.log("in Construrctor");
    initializeIcons();
    this.selection.getKey = (item => { return item["ID"]; });
    this.state = {
      primaryApprover: null,
      RoleReviewItems: [],
      showTcodePopup: false,
      showApprovalPopup: false
    };
  }
  public componentDidMount(): void {
    debugger;
    // this.fetchRoleReview();
    // this.fetchPrimaryApprover();
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



  public showPopup(item: RoleReviewItem) {
    this.fetchRoleToTransaction(item.Role_Name)
      .then((roleToTransactions) => {
        console.log(roleToTransactions);
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
    var tempArray = map(this.state.RoleReviewItems, (uaItem) => {
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
      RoleReviewItems: tempArray,
      popupValueApproval: null,
      popupValueComments: null,
      showApprovalPopup: false
    }));
  }
  // @autobind
  // public getApi(controller: string, query: string): Promise<any> {
  //   let url = this.props.webApiUrl + "/api/" + controller + "?" + query;
  //   let httpClientOptions: IHttpClientOptions = {
  //     credentials: "include",
  //   };
  //   return this.props.httpClient.get(url,    HttpClient.configurations.v1,httpClientOptions)
  //     .then((response: HttpClientResponse) => {
  //       return response.json();
  //     });
  // }

  // @autobind
  // public putApi(controller: string, entity: any): Promise<any> {
  //   let url = this.props.webApiUrl + "/api/" + controller + "/" + entity["ID"];
  //   let requestHeaders: Headers = new Headers();
  //   requestHeaders.append('Content-type', 'application/json');
  //   let httpClientOptions: IHttpClientOptions = {
  //     credentials: "include",
  //     headers: requestHeaders,
  //     method: "PUT",
  //     body: JSON.stringify(entity)
  //   };
  //   return this.props.httpClient.fetch(url, HttpClient.configurations.v1, httpClientOptions);
  // }
  @autobind
  public updateRoleReviewItems(items: RoleReviewItem[]): Promise<any> {
    let promises: Array<Promise<any>> = [];
    for (let item of items) {
      // promises.push(this.putApi(this.props.roleReviewController, item));
      let query = `${this.props.azureFunctionUrl}/api/RoleReviews/${item.ID}?&code=${this.props.accessCode}`;
      promises.push(this.props.httpClient.post(query, HttpClient.configurations.v1, {
        credentials: "include", referrerPolicy: "unsafe-url", body: item, method: "PUT"
      }));
    }
    return Promise.all(promises);
  }

  @autobind
  public setComplete(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let updatedApprover = this.state.primaryApprover;
    updatedApprover.Completed = "Yes";
    // this.putApi(this.props.primaryApproverController, updatedApprover)
    //   .then(() => {
    //     this.setState((current) => ({ ...current, primaryApprover: updatedApprover }));
    //   })
    //   .catch((err) => {
    //     console.log(err);
    //     alert("An error occurred saving the primary approver record");
    //   });
  }
  @autobind
  public unsetComplete(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let updatedApprover = this.state.primaryApprover;
    updatedApprover.Completed = "";
    // this.putApi(this.props.primaryApproverController, updatedApprover)
    //   .then(() => {
    //     this.setState((current) => ({ ...current, primaryApprover: updatedApprover }));
    //   })
    //   .catch((err) => {
    //     console.log(err);
    //     alert("An error occurred saving the primary approver record")
    //   });

  }
  @autobind
  public save(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    this.updateRoleReviewItems(this.state.RoleReviewItems).then(() => {
      var tempArray = map(this.state.RoleReviewItems, (rr) => {
        return { ...rr, hasBeenUpdated: false };
      });
      this.setState((current) => ({ ...current, RoleReviewItems: tempArray }));
      alert("Saved");
    }).catch((err) => {
      debugger;
      alert(err);
    });
  }
  //@autobind
  // public addApprover(approver: any): Promise<HttpClientResponse> {

  //   let requestHeaders: Headers = new Headers();
  //   requestHeaders.append('Content-type', 'application/json');
  //   requestHeaders.append('Cache-Control', 'no-cache');

  //   let httpClientOptions: IHttpClientOptions = {
  //     credentials: "include",
  //     body: JSON.stringify(approver)
  //   };

  //   //let url=`${this.props.webApiUrl}/api/${this.props.primaryApproverController}?$filter=tolower(ApproverEmail) eq '${approverEmail.toLowerCase()}'`;
  //   let url = `${this.props.webApiUrl}/api/${this.props.primaryApproverController}`;
  //   console.log(url);
  //   return this.props.httpClient.post(url,
  //     HttpClient.configurations.v1,
  //     httpClientOptions);
  // }

  @autobind
  public fetchPrimaryApprover(): Promise<any> {
    //let query = "$filter=tolower(ApproverEmail) eq '" + this.props.user.email.toLowerCase() + "'";
    let query = `${this.props.azureFunctionUrl}/api/PrimaryApprovers/${this.props.user.email}?&code=${this.props.accessCode}`;

    return this.props.httpClient.fetch(query, HttpClient.configurations.v1, {
      credentials: "include", referrerPolicy: "unsafe-url"
    })
      .then((response: HttpClientResponse) => {
        debugger;
        response.json()
          .then((appr) => {
            this.setState((current) => ({ ...current, primaryApprover: appr[0] }));
          })
          .catch((err) => {
            debugger
          })
      }).catch(e => {
        console.log(e);
        debugger;
        alert("There was an error fetching Primary approvers");
      });

  }
  @autobind
  public fetchRoleReview(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    //let query = "$filter=tolower(ApproverEmail) eq '" + this.props.user.email.toLowerCase() + "'";
    let query = `${this.props.azureFunctionUrl}/api/RoleReviews/${this.props.user.email}?&code=${this.props.accessCode}`;

    this.props.httpClient.fetch(query, HttpClient.configurations.v1, {
      credentials: "include", referrerPolicy: "unsafe-url"
    })

      .then((response: HttpClientResponse) => {
        response.json().then((rolereviews) => {
          this.setState((current) => ({ ...current, RoleReviewItems: rolereviews }));
        })
      })
      .catch((err) => {
        debugger;
      })
      .catch(err => {
        console.log(err);
        alert("There was an error fetching Role Review Items");
      });
  }
  @autobind
  public fetchRoleToTransaction(RoleName: string): Promise<any> {
    console.log(RoleName);
    let r2 = RoleName.replace(/\//g, "~");
    //let query = "$filter=Role_Name eq '" + RoleName + "'";
    let query = `${this.props.azureFunctionUrl}/api/RoleToTransactions/${r2}?&code=${this.props.accessCode}`;
    debugger;
    return this.props.httpClient.fetch(query, HttpClient.configurations.v1, {
      credentials: "include", referrerPolicy: "unsafe-url"
    })
      .then((response: HttpClientResponse) => {
        debugger;
        return response.json();
      })
      .catch((err) => {
        debugger;
      })


  }

  /**
   * this function gets called after the iframe has connected to the webapi.
   * After this we can make calls to the web api passing the credentials
   */
  public frameLoaded() {

    this.fetchRoleReview();
    this.fetchPrimaryApprover();
  }

  public RenderApproval(item?: RoleReviewItem, index?: number, column?: IColumn): JSX.Element {
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
            let tempRoleToTCodeReview = this.state.RoleReviewItems;
            let rtc = find(tempRoleToTCodeReview, (x) => {
              return x.ID === item.ID;
            });
            rtc.Approval = option.key as string;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, RoleReviewItems: tempRoleToTCodeReview, changesHaveBeenMade: true }));
          }}
        />
      );
    }
  }
  public RenderComments(item?: RoleReviewItem, index?: number, column?: IColumn): JSX.Element {
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
            let tempRoleToTCodeReview = this.state.RoleReviewItems;
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
  public render(): React.ReactElement<IRoleToTCodeProps> {
    console.log(this.props.azureFunctionUrl);
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

          if (!this.selection.count || this.selection.count < this.state.RoleReviewItems.length) {
            this.setState((current) => ({
              ...current,
              showApprovalPopup: true,
              changeSelected: false // change UNSELECTED Items
            }));
          }
        },

      },

      {
        key: "Undo", name: "Undo", icon: "Undo", onClick: this.fetchRoleReview,
        disabled: !(filter(this.state.RoleReviewItems, (rr) => {
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
        (filter(this.state.RoleReviewItems, (rr) => { return rr.Approval === "3"; }).length > 0) // "3" is the initial state after larry uploads the access db
      },
      {
        key: "UnDone", name: "UnComplete", icon: "Completed", onClick: this.unsetComplete
      }

    ];
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {

        key: "Save", name: "Save", icon: "Save", onClick: this.save,
        disabled: !(this.state.primaryApprover) || !(filter(this.state.RoleReviewItems, (rr) => { return rr.hasBeenUpdated; }).length > 0)
        || this.state.primaryApprover.Completed === "Yes"

      }
    ];


    return (
      <div className={styles.roleToTCode}>
        <iframe src={this.props.azureFunctionUrl} onLoad={this.frameLoaded.bind(this)} />

        <Dialog isBlocking={true}
          hidden={!this.state.showApprovalPopup}
          onDismiss={(e) => { this.setState((current) => ({ ...current, showApprovalPopup: false })); }}
          dialogContentProps={{
            type: DialogType.close,
            title: this.state.changeSelected
              ? `Update ${this.selection.count} Selected Items`
              : this.selection.count
                ? `Update ${this.state.RoleReviewItems.length - this.selection.count} Unselected Items`
                : `Update ${this.state.RoleReviewItems.length} Unselected Items`,
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
          items={this.state.RoleReviewItems}
          selectionMode={SelectionMode.multiple}
          selection={this.selection}
          key="ID"
          columns={[

            {
              key: "title", name: "Role Name",
              fieldName: "Role_Name", minWidth: 300, maxWidth: 300,
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
                fieldName: "Role", minWidth: 250, maxWidth: 250,
              },
              {
                key: "Comments", name: "TCode",
                isResizable: true,
                fieldName: "TCode", minWidth: 150, maxWidth: 150,

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
