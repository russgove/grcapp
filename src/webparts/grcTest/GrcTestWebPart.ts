import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "GrcTestWebPartStrings";
import GrcTest from "./components/GrcTest";
import { IGrcTestProps } from "./components/IGrcTestProps";
import pnp, { EmailProperties, Items, Item } from "sp-pnp-js";
import PrimaryApproverList from "../../dataModel/PrimaryApproverList";
import RoleReview from "../../dataModel/RoleReview";
import RoleToTransaction from "../../dataModel/RoleToTransaction";
import { find, filter } from "lodash";
export interface IGrcTestWebPartProps {
  primaryApproverListName: string;
  roleToTCodeReviewListName: string;
  roleToTransactionListName: string;
}

export default class GrcTestWebPart extends BaseClientSideWebPart<IGrcTestWebPartProps> {
  private primaryApproverLists: Array<PrimaryApproverList>;
  private roleReviews: Array<RoleReview>;
  private reactElement: React.ReactElement<IGrcTestProps>;
  private formComponent: GrcTest;

  private async fetchRoleReviews(): Promise<Array<RoleReview>> {
    debugger;
    let userId = this.context.pageContext.legacyPageContext.userId;
    let select = `Id,GRCRoleName,GRCApproverId,GRCApprover/Title,
    GRCApproval,GRCApprovedById, GRCDateReview, 
    GRCComments, GRCRemediation`;
    let expands = "GRCApprover";

    return pnp.sp.web.lists.getByTitle(this.properties.roleToTCodeReviewListName).items
      .select(select)
      .expand(expands)
      .filter('GRCApproverId eq ' + userId)
      .getAs<Array<RoleReview>>();
  }
  public async onInit(): Promise<void> {
    await super.onInit().then(() => {
      pnp.setup({
        spfxContext: this.context,
      });
      return;
    });
    let userId = this.context.pageContext.legacyPageContext.userId;
    // this.testmethod();
    let expands = "GRCApprover";
    let select = "Id,GRCCompleted,GRCApprover,GRCApproverId,GRCApprover/Title";
    await pnp.sp.web.lists.getByTitle(this.properties.primaryApproverListName).items
      .select(select)
      .expand(expands)
      .filter('GRCApproverId eq ' + userId)
      .getAs<Array<PrimaryApproverList>>().then((result) => {
        this.primaryApproverLists = result;

      }).catch((err) => {
        console.error(err.data.responseBody["odata.error"].message.value);
        alert(err.data.responseBody["odata.error"].message.value);
        debugger;
      });



  }
  public save(roleToTCodeReviews: Array<RoleReview>): Promise<any> {
    debugger;
    let itemsToSave = filter(roleToTCodeReviews, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = pnp.sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
      pnp.sp.web.lists.getByTitle(this.properties.roleToTCodeReviewListName)
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

  public getRoleToTransaction(roleName: string): Promise<RoleToTransaction> {
    debugger;


    return pnp.sp.web.lists.getByTitle(this.properties.roleToTransactionListName)
      .items.filter(`GRCRoleName eq '${roleName}'`).getAs<RoleToTransaction>();


  }
  public setComplete(primaryApproverList: PrimaryApproverList): Promise<any> {
    debugger;

    let userId = this.context.pageContext.legacyPageContext.userId;
    return pnp.sp.web.lists.getByTitle(this.properties.primaryApproverListName)
      .items.getById(primaryApproverList.Id).update({ "GRCCompleted": "Yes" }).then(() => {
        let newProps = this.reactElement.props;
        newProps.primaryApproverList[0].GRCCompleted = "Yes";
        this.reactElement.props = newProps;
        this.formComponent.forceUpdate();

      });

  }
  public render(): void {
    this.reactElement = React.createElement(
      GrcTest,
      {
        primaryApproverList: this.primaryApproverLists,

        save: this.save.bind(this),
        getRoleToTransaction: this.getRoleToTransaction.bind(this),
        setComplete: this.setComplete.bind(this),
        fetchRoleReviews: this.fetchRoleReviews.bind(this),
        domElement: this.domElement
      }
    );


    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as GrcTest;

  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
