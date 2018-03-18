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
import * as strings from 'HighRiskWebPartStrings';
import HighRisk from './components/HighRisk';
import { IHighRiskProps } from './components/IHighRiskProps';
import { PropertyFieldCodeEditor,PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import pnp from "sp-pnp-js";
import { find, filter } from "lodash";
import { HighRiskItem, PrimaryApproverItem ,RoleToTransaction} from "./dataModel";

export interface IHighRiskWebPartProps {
  primaryApproversListName: string;
  highRiskListName: string;
  roleToTransactionListName:string;
  effectiveLabel:string;
  continuesLabel:string;
  correctPersonLabel:string;
}

export default class HighRiskWebPart extends BaseClientSideWebPart<IHighRiskWebPartProps> {
  private primaryApproverLists: Array<any>;
  private highRisks: Array<any>;
  private reactElement: React.ReactElement<IHighRiskProps>;
  private formComponent: HighRisk;
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
  private async fetchHighRisks(): Promise<Array<HighRisk>> {
    debugger;
    let userId = this.context.pageContext.legacyPageContext.userId;
    let select = `Id,GRCRole,GRCRoleName,GRCApproverId,GRCApprover/Title,
    GRCApproval,GRCApprovedById, GRCDateReview, GRCUserId,GRCUserFullName,
    GRCComments, GRCRemediation`;
    let expands = "GRCApprover";

    return pnp.sp.web.lists.getByTitle(this.properties.highRiskListName).items
      .select(select)
      .expand(expands)
      .filter('GRCApproverId eq ' + userId)
      .getAs<Array<HighRisk>>();
  }

  public save(HighRisks: Array<HighRiskItem>): Promise<any> {
    debugger;
    let itemsToSave = filter(HighRisks, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = pnp.sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
      pnp.sp.web.lists.getByTitle(this.properties.highRiskListName)
        .items.getById(item.Id).inBatch(batch).update({ "GRCApproval": item.Approval, "GRCComments": item.Comments })
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


    return pnp.sp.web.lists.getByTitle(this.properties.roleToTransactionListName)
      .items.filter(`GRCRole eq '${role}'`).getAs<RoleToTransaction>();


  }
  public setComplete(primaryApproverList: any): Promise<any> {
    debugger;

    let userId = this.context.pageContext.legacyPageContext.userId;
    return pnp.sp.web.lists.getByTitle(this.properties.primaryApproversListName)
      .items.getById(primaryApproverList.Id).update({ "GRCCompleted": "Yes" }).then(() => {
        let newProps = this.reactElement.props;
        newProps.primaryApprover[0].Completed = "Yes";
        this.reactElement.props = newProps;
        this.formComponent.forceUpdate();

      });

  }
  public render(): void {
    const element: React.ReactElement<IHighRiskProps > = React.createElement(
      HighRisk,
      {
        primaryApprover: this.primaryApproverLists,
        save: this.save.bind(this),
        getRoleToTransaction: this.getRoleToTransaction.bind(this),
        setComplete: this.setComplete.bind(this),
        fetchHighRisks: this.fetchHighRisks.bind(this),
        domElement: this.domElement,
        
      }
    );

    ReactDom.render(element, this.domElement);
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
