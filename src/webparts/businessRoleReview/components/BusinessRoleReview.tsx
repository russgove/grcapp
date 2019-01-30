import * as React from 'react';
import styles from './BusinessRoleReview.module.scss';
import { IBusinessRoleReviewProps } from './IBusinessRoleReviewProps';
import { IBusinessRoleReviewState } from './IBusinessRoleReviewState';
import { escape } from '@microsoft/sp-lodash-subset';
import { BusinessRoleReviewItem, PrimaryApproverItem } from "../dataModel";
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import {
  DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, Selection,
  ColumnActionsMode
} from "office-ui-fabric-react/lib/DetailsList";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import { Modal, IModalProps } from "office-ui-fabric-react/lib/Modal";
import { Panel, IPanelProps, PanelType } from "office-ui-fabric-react/lib/Panel";
import { CommandBar } from "office-ui-fabric-react/lib/CommandBar";
import { Label } from "office-ui-fabric-react/lib/Label";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";

import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton, IconButton } from "office-ui-fabric-react/lib/Button";
import { Dialog, DialogType, DialogContent, DialogFooter } from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { find, map, clone, filter } from "lodash";
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { autobind } from "office-ui-fabric-react/lib/Utilities";
import { Overlay } from "office-ui-fabric-react/lib/Overlay";
import { Spinner ,SpinnerSize} from "office-ui-fabric-react/lib/Spinner";



export default class businessbusinessRoleReviewItems extends React.Component<IBusinessRoleReviewProps, IBusinessRoleReviewState> {
  private selection: Selection = new Selection();
  public constructor(props: IBusinessRoleReviewProps) {
    super(props);
    console.log("in Construrctor");
    initializeIcons();
    this.selection.getKey = (item => { return item["Id"]; });
    this.state = {
      primaryApprover: null,
      businessRoleReviewItems: [],

      showPopup: false,
      showOverlay: true,
      overlayMessage: "Loading ..."
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
  @autobind
  public fetchBusinessRoleReview(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): Promise<any> {

    //let query = "$filter=tolower(ApproverEmail) eq '" + this.props.user.email.toLowerCase() + "'";
    let query = `${this.props.azureFunctionUrl}/api/${this.props.businessRoleOwnersPath}/${this.props.user.email}?code=${this.props.accessCode}`;

    return this.props.httpClient.fetch(query, HttpClient.configurations.v1, { credentials: "include", referrerPolicy: "unsafe-url" })
      .then((response: HttpClientResponse) => {
        return response.json().then((rolereviews) => {
          this.setState((current) => (
            { ...current, businessRoleReviewItems: rolereviews }
          )
          );
        });
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
  public fetchPrimaryApprover(): Promise<any> {

    //let query = "$filter=tolower(ApproverEmail) eq '" + this.props.user.email.toLowerCase() + "'";
    let query = `${this.props.azureFunctionUrl}/api/${this.props.primaryApproversPath}/${this.props.user.email}?code=${this.props.accessCode}`;

    return this.props.httpClient.fetch(query, HttpClient.configurations.v1, {
      credentials: "include", referrerPolicy: "unsafe-url"
    })
      .then((response: HttpClientResponse) => {
        if (response.status !== 200) {
          alert(`No Primary Approver record found for ${this.props.user.email}. Please contact the system adminsitrator.`);
          this.setState((current) => ({ ...current, primaryApprover: null }));
          return;
        }
        return response.json()
          .then((appr) => {
            if (appr.length === 0) {
              alert(`No Primary Approver record found for ${this.props.user.email}. Please contact the system adminsitrator.`);

            }
            if (appr.length > 1) {
              alert(`Multiple  Primary Approver records found for ${this.props.user.email}. Please contact the system adminsitrator.`);

            }
            this.setState((current) => ({ ...current, primaryApprover: appr[0] }));
          })
          .catch((err) => {
            debugger;
            alert(err);
          });
      }).catch(e => {
        console.log(e);
        debugger;
        alert("There was an error fetching Primary approvers");
      });

  }
  public componentDidMount(): void {

    Promise.all([
      this.fetchBusinessRoleReview(),
      this.fetchPrimaryApprover()])
      .then((x) => {

        this.setState((current) => ({ ...current, showOverlay: false, overlayMessage: "" }));
      }
      );
  }

  public showUpdatedSelectedPopup(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {

    if (this.selection.count > 0) {
      this.setState((current) => ({ ...current, showPopup: true }));
    }
  }
  /**
   * 
   * This method gets called by the popup window to update all the selected items
   */
  @autobind
  public updateSelected(ev?: React.MouseEvent<HTMLElement>, item?: IContextualMenuItem): void {
    debugger;
    var tempArray = map(this.state.businessRoleReviewItems, (brrItem: BusinessRoleReviewItem) => {
      debugger;
      if (this.selection.isKeySelected(brrItem.Id.toString()) === this.state.changeSelected) {
        debugger;
        return {
          ...brrItem,
          Approval: this.state.popupValueApproval ? this.state.popupValueApproval : brrItem.Approval,
          Comments: this.state.popupValueComments ? this.state.popupValueComments : brrItem.Comments,
          hasBeenUpdated: true,
        };
      }
      else {
        return {
          ...brrItem
        };
      }
    });

    this.setState((current) => ({
      ...current,
      businessRoleReviewItems: tempArray,
      popupValueApproval: null,
      popupValueComments: null,
      showPopup: false
    }));
  }


  @autobind
  public setComplete(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let updatedApprover = this.state.primaryApprover;
    updatedApprover.Completed = "Yes";
    let query = `${this.props.azureFunctionUrl}/api/${this.props.primaryApproversPath}/${updatedApprover.Id}?code=${this.props.accessCode}`;
    this.props.httpClient.fetch(query, HttpClient.configurations.v1, {
      body: JSON.stringify(updatedApprover), method: "PUT", mode: "cors", referrerPolicy: "unsafe-url"
    })
      .then(() => {

        this.setState((current) => ({ ...current, primaryApprover: updatedApprover }));
        alert("Completed");
      })
      .catch((err) => {
        debugger;
        console.log(err);
        alert("An error occurred saving the primary approver record");
      });
  }
  @autobind
  public unsetComplete(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    let updatedApprover = this.state.primaryApprover;
    updatedApprover.Completed = "";
    let query = `${this.props.azureFunctionUrl}/api/${this.props.primaryApproversPath}/${updatedApprover.Id}?code=${this.props.accessCode}`;
    this.props.httpClient.fetch(query, HttpClient.configurations.v1, {
      body: JSON.stringify(updatedApprover), method: "PUT", mode: "cors", referrerPolicy: "unsafe-url"
    })
      .then(() => {
        this.setState((current) => ({ ...current, primaryApprover: updatedApprover }));
      })
      .catch((err) => {
        debugger;
        console.log(err);
        alert("An error occurred saving the primary approver record");
      });

  }
  @autobind
  public updateBusinessRoleReviewItems(items: BusinessRoleReviewItem[]): Promise<any> {
    var updatedItems = filter(this.state.businessRoleReviewItems, (brr) => {
      return brr.hasBeenUpdated;
    });
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-Type', 'application/json');// if you dont do this you get an error No Media type is available to read an object of type ...
    let query = `${this.props.azureFunctionUrl}/api/${this.props.businessRoleOwnersPath}?code=${this.props.accessCode}`;
    return this.props.httpClient.fetch(query, HttpClient.configurations.v1, {
      referrerPolicy: "unsafe-url",
      body: JSON.stringify(updatedItems), method: "PUT", mode: "cors",headers:requestHeaders
    });
  }

  @autobind
  public save(ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
    debugger;
    this.setState((current) => ({ ...current, showOverlay: true, overlayMessage: "Saving ..." }));
    this.updateBusinessRoleReviewItems(this.state.businessRoleReviewItems).then(() => {
      var tempArray = map(this.state.businessRoleReviewItems, (rr) => {
        return { ...rr, hasBeenUpdated: false };
      });
      this.setState((current) => ({ ...current, businessRoleReviewItems: tempArray, showOverlay: false }));
      alert("Saved");
    }).catch((err) => {
      debugger;
      this.setState((current) => ({ ...current, showOverlay: false }));
      alert(err);
    });
  }
  public RenderComments(item?: BusinessRoleReviewItem, index?: number, column?: IColumn): JSX.Element {
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
            let items = this.state.businessRoleReviewItems;
            let rtc = find(items, (x) => {
              return x.Id === item.Id;
            });
            rtc.Comments = newValue;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, businessRoleReviewItems: items, changesHaveBeenMade: true }));
          }}
        >
        </TextField>
      );
    }
  }
  public RenderApproval(item?: BusinessRoleReviewItem, index?: number, column?: IColumn): JSX.Element {
    let options = [
      { key: "1", text: "Yes" },
      { key: "2", text: "No" },

    ];
    if (this.state.primaryApprover.Completed === "Yes") {
      return (
        <div>
          {find(options, (o) => { return o.key === item[column.fieldName]; }).text}
        </div>
      );
    }
    else {
      return (
        <ChoiceGroup label=""
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
          selectedKey={item[column.fieldName]}

          onChanged={(option: IChoiceGroupOption, event: any) => {
            let tempTable = this.state.businessRoleReviewItems;

            let rtc = find(tempTable, (x) => {
              return x.Id === item.Id;
            });
            rtc[column.fieldName] = option.key as string;
            rtc.hasBeenUpdated = true;
            this.setState((current) => ({ ...current, businessRoleReviewItems: tempTable, changesHaveBeenMade: true }));
          }}
        />
        // <Dropdown options={options}
        //   selectedKey={item[column.fieldName]}
        //   onChanged={(option: IDropdownOption, idx?: number) => {
        //     let tempTable = this.state.businessbusinessRoleReviewItems;

        //     let rtc = find(tempTable, (x) => {
        //       return x.Id === item.Id;
        //     });
        //     rtc[column.fieldName] = option.key as string;
        //     rtc.hasBeenUpdated = true;
        //     this.setState((current) => ({ ...current, mitigatingControls: tempTable, changesHaveBeenMade: true }));
        //   }}
        // >
        // </Dropdown>
      );
    }
  }
  private areAllQuestionsAnswered(): boolean {
    for (let brr of this.state.businessRoleReviewItems) {

      if (brr.Approval === "3") {
        return false;
      }
      if (!brr.Comments) {
        return false;
      }
      if (brr.Comments.length === 0) {
        return false;
      }

    }

    return true;
  }
  public render(): React.ReactElement<IBusinessRoleReviewProps> {
debugger;
   const hasUnsavedChanges:boolean=(filter(this.state.businessRoleReviewItems, (rr) => { return rr.hasBeenUpdated; }).length > 0);
   const selectedItemCount:number=this.selection.getSelectedCount();

    let itemsNonFocusable: IContextualMenuItem[] = [
      {
        key: "Update Selected",
        name: "Update Selected",
        icon: "TriggerApproval",
        onClick: (e) => {
          if (this.selection.count > 0) {
            this.setState((current) => ({
              ...current,
              showPopup: true,
              changeSelected: true
            }));
          }
        },
        disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes" || this.selection.getSelectedCount() <1,

      },
      {
        key: "Update Unselected",
        name: "Update Unselected",
        icon: "TriggerAuto",
        disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes" || this.selection.getSelectedCount() <this.state.businessRoleReviewItems.length,
        onClick: (e) => {
          debugger;
          if (!this.selection.count || this.selection.getSelectedCount() < this.state.businessRoleReviewItems.length+1) {
            this.setState((current) => ({
              ...current,
              showPopup: true,
              changeSelected: false // change UNSELECTED Items
            }));
          }
        },

      },

      {
        key: "Undo", name: "Undo", icon: "Undo", onClick: this.fetchBusinessRoleReview,
        disabled: !(filter(this.state.businessRoleReviewItems, (rr) => {
          return rr
            .hasBeenUpdated;
        }).length > 0)
      },
      { // if the item has been comleted OR there are items with noo approvasl, diable
        key: "Done", name: "Task Complete", icon: "Completed", onClick: this.setComplete,
        disabled: !(this.state.primaryApprover) || this.state.primaryApprover.Completed === "Yes" || !(this.areAllQuestionsAnswered()) || hasUnsavedChanges

      }

    ];
    if (this.props.enableUncomplete) {
      itemsNonFocusable.push({
        key: "UnDone", name: "UnComplete", icon: "Completed", onClick: this.unsetComplete ,
        disabled: !(this.state.primaryApprover) ||  this.state.primaryApprover.Completed !== "Yes"
      });
    }
    let farItemsNonFocusable: IContextualMenuItem[] = [
      {

        key: "Save", name: "Save", icon: "Save", onClick: this.save,
        disabled: !hasUnsavedChanges
        || this.state.primaryApprover.Completed === "Yes"

      },
      {
        key: "helpLinks", name: "Help", icon: "help",
        items: map(this.props.helpLinks, (hl): IContextualMenuItem => {
          debugger;
          return {
            key: hl.Id.toString(), // this is the id of the listitem
            href: hl.Url.Url,
            title: hl.Url.Description,
            icon: hl.IconName,
            name: hl.Title,
            target: hl.Target

          };
        })
      }
    ];


    return (
      <div className={styles.businessRoleReview}>
      <Label>{selectedItemCount}</Label>
        <Dialog isBlocking={true}
          hidden={!this.state.showPopup}
          onDismiss={(e) => { this.setState((current) => ({ ...current, showPopup: false })); }}
          dialogContentProps={{
            type: DialogType.close,
            title: this.state.changeSelected
              ? `Update ${this.selection.count} Selected Items`
              : this.selection.count
                ? `Update ${this.state.businessRoleReviewItems.length - this.selection.count} Unselected Items`
                : `Update ${this.state.businessRoleReviewItems.length} Unselected Items`,
            subText: 'All selected items will be updated with the following values'
          }} >
          <ChoiceGroup label="Approval Decision" 
            options={[
              {
                key: '1',
                text: 'Yes',
                disabled: !(this.state.primaryApprover) ||this.state.primaryApprover.Completed === "Yes" 
              },
              {
                key: '2',
                text: 'No',
                disabled: !(this.state.primaryApprover) ||this.state.primaryApprover.Completed === "Yes" 
              },
            ]}
            selectedKey={this.state.popupValueApproval}
            onChange={(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption) => {

              this.setState((current) => ({ ...current, popupValueApproval: option.key }));
            }}
          />

          <TextField label="Comments" 
          disabled= {!(this.state.primaryApprover) ||this.state.primaryApprover.Completed === "Yes" }
           onChanged={(e) => {

            this.setState((current) => ({ ...current, popupValueComments: e }));
          }}

          />



          <DialogFooter>
            <PrimaryButton text='Save' onClick={this.updateSelected.bind(this)} />
            <DefaultButton text='Cancel' onClick={(e) => {
              this.setState((current) => ({ ...current, showPopup: false }));
            }} />
          </DialogFooter>
        </Dialog>
        <CommandBar
          isSearchBoxVisible={false}
          items={itemsNonFocusable}
          farItems={farItemsNonFocusable}

        />
        <DetailsList
          items={this.state.businessRoleReviewItems}
          selectionMode={SelectionMode.multiple}
          selection={this.selection}
          key="Risk_x0020_ID"
          layoutMode={DetailsListLayoutMode.justified}
          columns={[
            {
              key: "Role_x0020_Name", name: "Role Name / Composite Role",
              fieldName: "Role_x0020_Name",
              minWidth: this.props.roleNameWidth,
              isResizable: true,
              onRender: (item?: BusinessRoleReviewItem, index?: number, column?: IColumn) => {
                return (
                  <div  >
                    {item.RoleName}
                    <br />
                    {item.CompositeRole}
                  </div>
                );
              },
            },

            // {
            //   key: "Composite_x0020_Role", name: "Composite Role",
            //   fieldName: "Composite_x0020_Role", minWidth: this.props.compositeRoleWidth,
            // },
            {
              key: "Approver", name: "Approver ID/Name ",
              fieldName: "Approver",
              minWidth: this.props.approverWidth,
              isResizable: true,
              onRender: (item?: BusinessRoleReviewItem, index?: number, column?: IColumn) => {
                return (
                  <div  >
                    {item.Approver}
                    <br />
                    {item.ApproverName}
                  </div>
                );
              },
            },
            {
              key: "AlternateApprover", name: "Alt. Apprv. ID/Name",
              fieldName: "Composite_x0020_Role",
              minWidth: this.props.altApproverWidth,
              isResizable: true,
              onRender: (item?: BusinessRoleReviewItem, index?: number, column?: IColumn) => {
                return (
                  <div  >
                    {item.AltApprv}
                    <br />
                    {item.AlternateApprover}
                  </div>
                );
              },
            },

            {
              key: "Approval", name: "Approval Decision",
              fieldName: "Approval",
              minWidth: this.props.approvalDecisionWidth,
              isResizable: true,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderApproval(item, index, column);
              },

            },
            {
              key: "Comments", name: "Comments",
              fieldName: "Comments",
              minWidth: this.props.commentsWidth,
              isResizable: true,
              onRender: (item?: any, index?: number, column?: IColumn) => {
                return this.RenderComments(item, index, column);
              },
            },


          ]}
        />

{this.state.showOverlay && (
          <Overlay >



            <br /><br /><br /><br /><br /><br /><br />

            <Spinner size={SpinnerSize.large} label={this.state.overlayMessage} ariaLive="assertive" />


          </Overlay>
        )}

      </div>
    );
  }
}
