import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import{
  Environment,
  EnvironmentType
  } from '@microsoft/sp-core-library';
import * as strings from 'BusinessRoleReviewWebPartStrings';
import BusinessRoleReview from './components/BusinessRoleReview';
import { IBusinessRoleReviewProps } from './components/IBusinessRoleReviewProps';
import { PropertyFieldCodeEditor,PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import pnp from "sp-pnp-js";
import { find, filter } from "lodash";
import { BusinessRoleReviewItem, PrimaryApproverItem } from "./dataModel";

export interface IBusinessRoleReviewWebPartProps {
  primaryApproversListName: string;
  businessRoleReviewListName: string;
  effectiveLabel:string;
  continuesLabel:string;
  correctPersonLabel:string;
}

export default class BusinessRoleReviewWebPart extends BaseClientSideWebPart<IBusinessRoleReviewWebPartProps> {
  private primaryApproverLists: Array<any>;
  private highRisks: Array<any>;
  private reactElement: React.ReactElement<IBusinessRoleReviewProps>;
  private formComponent: BusinessRoleReview;
  public async onInit(): Promise<void> {

    await super.onInit().then(() => {
      pnp.setup({
        spfxContext: this.context,
      });
      return;
    });
    let userId = this.context.pageContext.legacyPageContext.userId;
    // this.testmethod();
    let expands = "PrimaryApprover";
    let select = "Id,Completed,PrimaryApprover,PrimaryApproverId,PrimaryApprover/Title";
    return pnp.sp.web.lists.getByTitle(this.properties.primaryApproversListName).items
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

    return pnp.sp.web.lists.getByTitle(this.properties.businessRoleReviewListName).items
      .select(select)
      .expand(expands)
      .filter('PrimaryApproverId eq ' + userId)
      .getAs<Array<BusinessRoleReviewItem>>();
  }
  public setComplete(primaryApproverList: any): Promise<any> {

    let userId = this.context.pageContext.legacyPageContext.userId;
    return pnp.sp.web.lists.getByTitle(this.properties.primaryApproversListName)
      .items.getById(primaryApproverList.Id).update({ "Completed": "Yes" }).then(() => {
        let newProps = this.reactElement.props;
        newProps.primaryApprover[0].Completed = "Yes";
        this.reactElement.props = newProps;
        this.formComponent.forceUpdate();

      });

  }
  public save(MitigatingControls: Array<BusinessRoleReviewItem>): Promise<any> {
    let itemsToSave = filter(MitigatingControls, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = pnp.sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
      pnp.sp.web.lists.getByTitle(this.properties.businessRoleReviewListName)
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
        effectiveLabel: this.properties.effectiveLabel,
        continuesLabel: this.properties.continuesLabel,
        correctPersonLabel: this.properties.correctPersonLabel
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
                PropertyPaneTextField('description', {
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
