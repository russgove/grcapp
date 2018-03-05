import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MitigatingControlsWebPartStrings';
import MitigatingControls from './components/MitigatingControls';
import { IMitigatingControlsProps } from './components/IMitigatingControlsProps';
import pnp, { EmailProperties, Items, Item } from "sp-pnp-js";
import { find, filter } from "lodash";
//import fieldNames from "./fieldNames"
import { MitigatingControlsItem, PrimaryApproverItem } from "./dataModel";
export interface IMitigatingControlsWebPartProps {
  primaryApproversListName: string;
  mitigatingControlsListName: string
}

export default class MitigatingControlsWebPart extends BaseClientSideWebPart<IMitigatingControlsWebPartProps> {
  private primaryApproverLists: Array<any>;
  private highRisks: Array<any>;
  private reactElement: React.ReactElement<IMitigatingControlsProps>;
  private formComponent: MitigatingControls;
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
  private async fetchMitigatingControls(): Promise<any> {
  
    let userId = this.context.pageContext.legacyPageContext.userId;
    let select = `*,PrimaryApproverId,PrimaryApprover/Title`;
    let expands = "PrimaryApprover";

    return pnp.sp.web.lists.getByTitle(this.properties.mitigatingControlsListName).items
      .select(select)
      .expand(expands)
      .filter('PrimaryApproverId eq ' + userId)
      .getAs<Array<MitigatingControlsItem>>();
  }
  public render(): void {
    this.reactElement = React.createElement(
      MitigatingControls,
      {
        primaryApprover: this.primaryApproverLists,
        save: this.save.bind(this),
        fetchMitigatingControls: this.fetchMitigatingControls.bind(this),
        setComplete: this.setComplete.bind(this),
        domElement: this.domElement
      }
    );

    this.formComponent = ReactDom.render(this.reactElement, this.domElement) as MitigatingControls;
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
  public save(MitigatingControls: Array<MitigatingControlsItem>): Promise<any> {
  let itemsToSave = filter(MitigatingControls, (rtc) => { return rtc.hasBeenUpdated === true; });
    let batch = pnp.sp.createBatch();
    //let promises: Array<Promise<any>> = [];

    for (let item of itemsToSave) {
      pnp.sp.web.lists.getByTitle(this.properties.mitigatingControlsListName)
        .items.getById(item.Id).inBatch(batch).update({ 
          "Effective": item.Effective, 
          "Continues": item.Continues, 
          "Right_x0020_Monitor_x003f_": item.Right_x0020_Monitor_x003f_, 
          "Comments": item.Comments })
        .then((x) => {
      
        })
        .catch((err) => {
          debugger;
        });

    }
    return batch.execute();

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
                PropertyPaneTextField('mitigatingControlsListName', {
                  label: "High Risk with Mitigating Controls List"
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
