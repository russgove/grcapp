import * as React from 'react';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import styles from './HighRiskAdminWebpart.module.scss';
import { IHighRiskAdminWebpartProps } from './IHighRiskAdminWebpartProps';
import { IHighRiskAdminWebpartState } from './IHighRiskAdminWebpartState';
import { escape } from '@microsoft/sp-lodash-subset';
const parse = require('csv-parse');
import pnp, { TypedHash, ItemAddResult, ListAddResult, ContextInfo, Web, WebAddResult, List as PNPList } from "sp-pnp-js";
import { List } from "office-ui-fabric-react/lib/List";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";
import { Button, IconButton, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { find, clone, map } from "lodash";
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('sp-workflow');
import {CachedId,findId,uploadFile,esnureUsers,extractColumnHeaders} from "../../../utilities/Utilities";
export default class HighRiskAdminWebpart extends React.Component<IHighRiskAdminWebpartProps,IHighRiskAdminWebpartState> {
  private addMessage(message: string) {
    let messages = this.state.messages;
    var copy = map(this.state.messages, clone);
    copy.push(message);
    this.setState((current) => ({ ...current, messages: copy }));
  }
  public render(): React.ReactElement<IHighRiskAdminWebpartProps> {
    return (
      <div className={ styles.highRiskAdminWebpart }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
