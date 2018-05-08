import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import {
  SPHttpClient, 
  SPHttpClientConfiguration, 
  SPHttpClientResponse, 
  ODataVersion, 
  ISPHttpClientConfiguration
} from '@microsoft/sp-http';
import {
  IODataWeb,
} from '@microsoft/sp-odata-types';
import { Dialog } from '@microsoft/sp-dialog';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './CommonHeaderFooterApplicationCustomizer.module.scss';

import * as strings from 'CommonHeaderFooterApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CommonHeaderFooterApplicationCustomizer';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICommonHeaderFooterApplicationCustomizerProperties {
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CommonHeaderFooterApplicationCustomizer
  extends BaseApplicationCustomizer<ICommonHeaderFooterApplicationCustomizerProperties> {
  
    private _TopPlaceHolder: PlaceholderContent | undefined;
    private _BottomPlaceHolder : PlaceholderContent | undefined;
    private _WebProperty : string;
    

  @override
  public onInit(): Promise<void> {
    console.log("Starting Application Customizer");
    
    //this.context.placeholderProvider.changedEvent.add(this, this._GetWebDetails);

    this._GetWebDetails();

    console.log("Starting Application Customizer");
    return Promise.resolve();
  }
  
  private _GetWebDetails(): Promise<string> {
    const spHTTPClient : SPHttpClient = this.context.spHttpClient;
    const currentWebURL : string = this.context.pageContext.web.absoluteUrl;

    spHTTPClient.get(`${currentWebURL}/_api/web/AllProperties?$select=cust_Site_Type`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        response.json().then((data) => {
            console.log(data.cust_x005f_Site_x005f_Type);
             this._WebProperty = data.cust_x005f_Site_x005f_Type;
        }).catch(() => {
          console.log("Unable to retrieve property !");
          this._WebProperty = "Classification yet to be defined by Admin";
        }).then(() => this._RenderPlaceHolders()).catch(() => {
          console.error();
        });
    });
    return Promise.resolve(this._WebProperty);
  }

  private _RenderPlaceHolders(): void {
    
    console.log("Available Placeholders: " + this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(", "));

    if(!this._BottomPlaceHolder){
      this._BottomPlaceHolder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        {onDispose: this._onDispose}
      );
    }

    if(!this._BottomPlaceHolder){
      console.log("Not Able to Fetch {0} Place Holder", PlaceholderName.Bottom);
    }

    if(!this._WebProperty){
      this._WebProperty = "Classification yet to be defined by Admin";
    }

    if(this._BottomPlaceHolder.domElement){
      this._BottomPlaceHolder.domElement.innerHTML=`<div class="${styles.app}">
      <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.bottom} ${styles.outerStruct}">
        <span style="margin-left:2%; display:inline-flex;">
          <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>&nbsp;${escape(this._WebProperty)}
        </span>
        <span style="margin-right:2%">
          &copy; ${escape(new Date().getFullYear().toString())}
        </span>
      </div>
    </div>`;
    }

  }

  private _onDispose(): void {
    console.log('Disposed custom top and bottom placeholders.');
  }
}
