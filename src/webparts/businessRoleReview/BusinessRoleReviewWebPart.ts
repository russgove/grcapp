import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,IPropertyPaneSliderProps
} from '@microsoft/sp-webpart-base';
import{
  Environment,
  EnvironmentType
  } from '@microsoft/sp-core-library';
import * as strings from 'BusinessRoleReviewWebPartStrings';
import BusinessRoleReview from './components/BusinessRoleReview';
import { IBusinessRoleReviewProps } from './components/IBusinessRoleReviewProps';
import { PropertyFieldCodeEditor,PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import {sp} from "@pnp/sp";
import { find, filter } from "lodash";
import { BusinessRoleReviewItem, PrimaryApproverItem,HelpLink } from "./dataModel";

export interface IBusinessRoleReviewWebPartProps {
  primaryApproversListName: string;
  businessRoleReviewListName: string;
  roleNameWidth:number;
 // compositeRoleWidth:number;
  approverWidth:number;
  altApproverWidth:number;
  approvalDecisionWidth:number;
  commentsWidth:number;
  helpLinksListName:string

}

export default class BusinessRoleReviewWebPart extends BaseClientSideWebPart<IBusinessRoleReviewWebPartProps> {
  private primaryApproverLists: Array<any>;
  private highRisks: Array<any>;
  private reactElement: React.ReactElement<IBusinessRoleReviewProps>;
  private formComponent: BusinessRoleReview;
  private helpLinks: Array<HelpLink>;
  public async onInit(): Promise<void> {

    await super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
      return;
    });
    await sp.site.rootWeb.lists.getByTitle(this.properties.helpLinksListName).items
    .filter("Audit eq'Mitigating Controls' or Audit eq 'All'")
    .get<Array<HelpLink>>().then((helps => {
      debugger;
      this.helpLinks = helps;
    })).catch((err) => {
      console.error(err);
      debugger;
      alert("There was an error fetching the helplinks");
      alert(err.data.responseBody["odata.error"].message.value);
    });
    let userId = this.context.pageContext.legacyPageContext.userId;
    // this.testmethod();
    let expands = "PrimaryApprover";
    let select = "Id,Completed,PrimaryApprover,PrimaryApproverId,PrimaryApprover/Title";
    return sp.web.lists.getByTitle(this.properties.primaryApproversListName).items
      .select(select)
      .expand(expands)
      .filter('PrimaryApproverId eq ' + userId)
      .get()
      .then((result) => {
        this.primaryApproverLists = result;

      }).catch((err) => {
      
        console.error(err.data.responseBody["odata.error"].message.value);
        alert(err.data.responseBody["odata.error"].message.value);
        debugger;
      });



  }
  private async fetchBusinessRoleReview(): Promise<any> {
  
    let userId = this.context.pageContext.legacyPageContext.userId;
    let select = `*,PrimaryApproverId,PrimaryApprover/Title`;
    let expands = "PrimaryApprover";

    return sp.web.lists.getByTitle(this.properties.businessRoleReviewListName).items
      .select(select)
      .expand(expands)
      .filter('PrimaryApproverId eq ' + userId)
      .get<Array<BusinessRoleReviewItem>>();
  }
  public setComplete(primaryApproverList: any): Promise<any> {

    let userId = this.context.pageContext.legacyPageContext.userId;
    return sp.web.lists.getByTitle(this.properties.primaryApproversListName)
      .items.getById(primaryApproverList.Id).update({ "Completed": "Yes" }).then(() => {
        let newProps = this.reactElement.props;
        newProps.primaryApprover[0].Completed = "Yes";
        this.reactElement.props = newProps;
        this.formComponent.forceUpdate();

      });

  }
  public save(MitigatingControls: Array<BusinessRoleReviewItem>): Promise<any> {
    let itemsToSave = filter(MitigatingControls, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
      sp.web.lists.getByTitle(this.properties.businessRoleReviewListName)
        .items.getById(item.Id).inBatch(batch).update({
          "Approval": item.Approval,
          "Comments": item.Comments
        })
        .then((x) => {

        })
        .catch((err) => {
          console.error(err);
          alert(err);
          debugger;
        });

    }
    return batch.execute();

  }
  public render(): void {
    this.reactElement = React.createElement(
      BusinessRoleReview,
      {
        primaryApprover: this.primaryApproverLists,
        save: this.save.bind(this),
        fetchBusinessRoleReview: this.fetchBusinessRoleReview.bind(this),
        setComplete: this.setComplete.bind(this),
        domElement: this.domElement,
        roleNameWidth:this.properties.roleNameWidth,
     //   compositeRoleWidth:this.properties.compositeRoleWidth,
        approverWidth:this.properties.approverWidth,
        altApproverWidth:this.properties.altApproverWidth,
        approvalDecisionWidth:this.properties.approvalDecisionWidth,
        commentsWidth:this.properties.commentsWidth,
        helpLinks: this.helpLinks

      }
    );

    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as BusinessRoleReview;

    if (Environment.type === EnvironmentType.ClassicSharePoint) {
      const buttons: NodeListOf<HTMLButtonElement> = this.domElement.getElementsByTagName('button');
      if (buttons && buttons.length) {
        for (let i: number = 0; i < buttons.length; i++) {
          if (buttons[i]) {
            /* tslint:disable */
            // Disable the button onclick postback
            buttons[i].onclick = function () { return false; };
            /* tslint:enable */
          }
        }
      }
    }
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
                PropertyPaneTextField('businessRoleReviewListName', {
                  label: "Business Role Owers List"
                }),
                PropertyPaneSlider("roleNameWidth",{
                  min:10,
                  max:1000,
                  label:"Width of Role Name column",
                  showValue:true
                }),
                // PropertyPaneSlider("compositeRoleWidth",{
                //   min:10,
                //   max:1000,
                //   label:"Width of Composite Role Name column",
                //   showValue:true
                // }),
                PropertyPaneSlider("approverWidth",{
                  min:10,
                  max:1000,
                  label:"Width of Approver column",
                  showValue:true
                }),
                PropertyPaneSlider("altApproverWidth",{
                  min:10,
                  max:1000,
                  label:"Width of Alternate Approver column",
                  showValue:true
                }),
                PropertyPaneSlider("approvalDecisionWidth",{
                  min:10,
                  max:1000,
                  label:"Width of Approval Decision column",
                  showValue:true
                }),
                PropertyPaneSlider("commentsWidth",{
                  min:10,
                  max:1000,
                  label:"Width of Comments column",
                  showValue:true
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
