import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'UserAccessWebPartStrings';
import UserAccess from './components/UserAccess';
import { IUserAccessProps } from './components/IUserAccessProps';
import { UserAccessItem, RoleToTransaction, PrimaryApproverItem } from "./datamodel";

import { find, filter } from "lodash";
import { sp, EmailProperties, Items, Item } from "@pnp/sp";

export interface IUserAccessWebPartProps {
  primaryApproversListName: string;
  highRiskListName: string;
  roleToTransactionListName: string;
}

export default class UserAccessWebPart extends BaseClientSideWebPart<IUserAccessWebPartProps> {
  private primaryApproverLists: Array<PrimaryApproverItem>;
  private highRisks: Array<UserAccessItem>;
  private reactElement: React.ReactElement<IUserAccessProps>;
  private formComponent: UserAccess;
  private async fetchUserAccess(): Promise<Array<UserAccessItem>> {
 
    let userId = this.context.pageContext.legacyPageContext.userId;
    let select = `Id,Role,Role_x0020_name,PrimaryApproverId,PrimaryApprover/Title,
    Approval, Date_x0020_Reviewed, User_x0020_ID,User_x0020_Full_x0020_Name,
    Comments, Remediation`;
    let expands = "PrimaryApprover";

    return sp.web.lists.getByTitle(this.properties.highRiskListName).items
      .select(select)
      .expand(expands)
      .filter('PrimaryApproverId eq ' + userId)
      .getAll();
  }
  public async onInit(): Promise<void> {
 
    await super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
      return;
    });
    let userId = this.context.pageContext.legacyPageContext.userId;
    // this.testmethod();
    let expands = "PrimaryApprover";
    let select = "Id,Completed,PrimaryApprover,PrimaryApproverId,PrimaryApprover/Title";
    return sp.web.lists.getByTitle(this.properties.primaryApproversListName).items
      .select(select)
      .expand(expands)
      .filter('PrimaryApproverId eq ' + userId)
      .get<Array<PrimaryApproverItem>>().then((result) => {
        this.primaryApproverLists = result;

      }).catch((err) => {
        console.error(err.data.responseBody["odata.error"].message.value);
        alert(err.data.responseBody["odata.error"].message.value);
        debugger;
      });



  }
  public save(HighRisks: Array<UserAccessItem>): Promise<any> {
  
    let itemsToSave = filter(HighRisks, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
      sp.web.lists.getByTitle(this.properties.highRiskListName)
        .items.getById(item.Id).inBatch(batch).update({ "Approval": item.Approval, "Comments": item.Comments })
        .then((x) => {
         
        })
        .catch((err) => {
          debugger;
        });

    }
   
    return batch.execute();

  }
  public getRoleToTransaction(role: string): Promise<RoleToTransaction[]> {
    let results: Promise<RoleToTransaction[]> =
      sp.web.lists.getByTitle(this.properties.roleToTransactionListName)
        .items.filter(`Composite_x0020_role eq '${role}'`).getAll();
    return results;

  }
  public setComplete(primaryApproverList: PrimaryApproverItem): Promise<any> {


    let userId = this.context.pageContext.legacyPageContext.userId;
    return sp.web.lists.getByTitle(this.properties.primaryApproversListName)
      .items.getById(primaryApproverList.Id).update({ "GRCCompleted": "Yes" }).then(() => {
        let newProps = this.reactElement.props;
        newProps.primaryApproverList[0].Completed = "Yes";
        this.reactElement.props = newProps;
        this.formComponent.forceUpdate();

      });

  }
  public render(): void {


    this.reactElement = React.createElement(
      UserAccess,
      {
        primaryApproverList: this.primaryApproverLists,
        save: this.save.bind(this),
        getRoleToTransaction: this.getRoleToTransaction.bind(this),
        setComplete: this.setComplete.bind(this),
        fetchUserAccess: this.fetchUserAccess.bind(this),
        domElement: this.domElement
      }
    );
  
    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as UserAccess;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('primaryApproversListName', {
                  label: "Primary Approvers List"
                }),
                PropertyPaneTextField('highRiskListName', {
                  label: "High Risk List"
                }),
                PropertyPaneTextField('roleToTransactionListName', {
                  label: "Role to Transaction List"
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
