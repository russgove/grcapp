import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HighRiskUserWebpartWebPartStrings';
import HighRiskUserWebpart from './components/HighRiskUserWebpart';
import { IHighRiskUserWebpartProps } from './components/IHighRiskUserWebpartProps';
import  { sp,EmailProperties, Items, Item } from "@pnp/sp";
import PrimaryApproverList from "../../dataModel/PrimaryApproverList";
import HighRisk from "../../dataModel/HighRisk";
import RoleToTransaction from "../../dataModel/RoleToTransaction";
import { find, filter } from "lodash";

export interface IHighRiskUserWebpartWebPartProps {
  primaryApproversListName: string;
  highRiskListName: string;
  roleToTransactionListName: string;
}

export default class HighRiskUserWebpartWebPart extends BaseClientSideWebPart<IHighRiskUserWebpartWebPartProps> {
  private primaryApproverLists: Array<PrimaryApproverList>;
  private highRisks: Array<HighRisk>;
  private reactElement: React.ReactElement<IHighRiskUserWebpartProps>;
  private formComponent: HighRiskUserWebpart;
  
  private async fetchHighRisks(): Promise<Array<HighRisk>> {
    debugger;
    let userId = this.context.pageContext.legacyPageContext.userId;
    let select = `Id,GRCRole,GRCRoleName,GRCApproverId,GRCApprover/Title,
    GRCApproval,GRCApprovedById, GRCDateReview, GRCUserId,GRCUserFullName,
    GRCComments, GRCRemediation`;
    let expands = "GRCApprover";

    return sp.web.lists.getByTitle(this.properties.highRiskListName).items
      .select(select)
      .expand(expands)
      .filter('GRCApproverId eq ' + userId)
      .get<Array<HighRisk>>();
  }
  public async onInit(): Promise<void> {
    debugger;
    await super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
      return;
    });
    let userId = this.context.pageContext.legacyPageContext.userId;
    // this.testmethod();
    let expands = "GRCApprover";
    let select = "Id,GRCCompleted,GRCApprover,GRCApproverId,GRCApprover/Title";
    return sp.web.lists.getByTitle(this.properties.primaryApproversListName).items
      .select(select)
      .expand(expands)
      .filter('GRCApproverId eq ' + userId)
      .get<Array<PrimaryApproverList>>().then((result) => {
        this.primaryApproverLists = result;

      }).catch((err) => {
        console.error(err.data.responseBody["odata.error"].message.value);
        alert(err.data.responseBody["odata.error"].message.value);
        debugger;
      });



  }
  public save(HighRisks: Array<HighRisk>): Promise<any> {
    debugger;
    let itemsToSave = filter(HighRisks, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
      sp.web.lists.getByTitle(this.properties.highRiskListName)
        .items.getById(item.Id).inBatch(batch).update({ "GRCApproval": item.GRCApproval, "GRCComments": item.GRCComments })
        .then((x) => {
          debugger;
        })
        .catch((err) => {
          debugger;
        });

    }
    debugger;
    return batch.execute();

  }
  public getRoleToTransaction(role: string): Promise<RoleToTransaction> {
    debugger;


    return sp.web.lists.getByTitle(this.properties.roleToTransactionListName)
      .items.filter(`GRCRole eq '${role}'`).get<RoleToTransaction>();


  }
  public setComplete(primaryApproverList: PrimaryApproverList): Promise<any> {
    debugger;

    let userId = this.context.pageContext.legacyPageContext.userId;
    return sp.web.lists.getByTitle(this.properties.primaryApproversListName)
      .items.getById(primaryApproverList.Id).update({ "GRCCompleted": "Yes" }).then(() => {
        let newProps = this.reactElement.props;
        newProps.primaryApproverList[0].GRCCompleted = "Yes";
        this.reactElement.props = newProps;
        this.formComponent.forceUpdate();

      });

  }
  public render(): void {
    debugger;
    this.reactElement = React.createElement(
      HighRiskUserWebpart,
      {
        primaryApproverList: this.primaryApproverLists,
        save: this.save.bind(this),
        getRoleToTransaction: this.getRoleToTransaction.bind(this),
        setComplete: this.setComplete.bind(this),
        fetchHighRisk: this.fetchHighRisks.bind(this),
        domElement: this.domElement
      }
    );
    debugger;
    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as HighRiskUserWebpart;
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
