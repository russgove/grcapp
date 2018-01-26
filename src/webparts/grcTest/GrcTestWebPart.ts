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
import RoleToTCodeReview from "../../dataModel/RoleToTCodeReview";
import { find, filter } from "lodash";
export interface IGrcTestWebPartProps {
  primaryApproverListName: string;
  roleToTCodeReviewListName: string;
  roleToTransactionListName: string;
}

export default class GrcTestWebPart extends BaseClientSideWebPart<IGrcTestWebPartProps> {
  private primaryApproverLists: Array<PrimaryApproverList>;
  private roleToTCodeReviews: Array<RoleToTCodeReview>;
  private async testmethod(): Promise<any> {
    class CustomListItem extends Item {
     public Id: number;
     public Title: string;
    }
    let listItems: Array<CustomListItem>;

    await pnp.sp.web.lists.getByTitle("CustomList").items.getAs<Array<CustomListItem>>()
      .then(response => {
        listItems = response;
      });

    for (let listItem of listItems) {
      listItem.update({ "Title": "NewTitle" });
    }
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
    let expands = "AssignedTo";
    let select = "Id,Approver,Approver_x0020_Name,Completed,AssignedTo,AssignedToId,AssignedTo/Title";
    await pnp.sp.web.lists.getByTitle(this.properties.primaryApproverListName).items
      .select(select)
      .expand(expands)
      .filter('AssignedToId eq ' + userId)
      .getAs<Array<PrimaryApproverList>>().then((result) => {
        this.primaryApproverLists = result;

      }).catch((err) => {
        debugger;
      });
    select = `Id,Role_x0020_Name,Approver,Approver_x0020_Name,
      Alt_x0020_Apprv, Alternate_x0020_Approver,
      Approval, Approved_x0020_By, Date_x0020_Reviewed, 
      Comments, Remediation`;
    await pnp.sp.web.lists.getByTitle(this.properties.roleToTCodeReviewListName).items
      .select(select)
      .filter(`Approver eq '${this.primaryApproverLists[0].Approver}'`)
      .getAs<Array<RoleToTCodeReview>>().then((result) => {
        this.roleToTCodeReviews = result;
      }).catch((err) => {
        console.error(err);
        alert(err.data.responseBody["odata.error"].message.value);
      });

  }
  public save(roleToTCodeReviews: Array<RoleToTCodeReview>): Promise<any> {
    debugger;
    let itemsToSave = filter(roleToTCodeReviews, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = pnp.sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
    pnp.sp.web.lists.getByTitle(this.properties.roleToTCodeReviewListName)
        .items.getById(item.Id).inBatch(batch).update({ "Approval": item.Approval })
        .then((x) => {
          debugger;
        })
        .catch((err) => {
          debugger;
        });
      
    }
    return batch.execute();

    // let batch = pnp.sp.createBatch();
    // let promises:Array<Promise<any>>=[];
    // for (let item of itemsToSave) {
    //   let promise=item.update({ "Approval": item.Approval })
    //     .then((x) => {
    //       debugger;
    //     })
    //     .catch((err) => {
    //       debugger;
    //     })
    //     promises.push(promise)
    // }
    // return Promise.all(promises);
  }
  public render(): void {
    const element: React.ReactElement<IGrcTestProps> = React.createElement(
      GrcTest,
      {
        primaryApproverList: this.primaryApproverLists,
        roleToTCodeReview: this.roleToTCodeReviews,
        save: this.save.bind(this)
      }
    );

    ReactDom.render(element, this.domElement);
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
