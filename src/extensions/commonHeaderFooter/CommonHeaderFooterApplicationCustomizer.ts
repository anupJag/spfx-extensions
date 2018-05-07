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
  IODataWeb
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
  Bottom:string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CommonHeaderFooterApplicationCustomizer
  extends BaseApplicationCustomizer<ICommonHeaderFooterApplicationCustomizerProperties> {
    private _TopPlaceHolder: PlaceholderContent | undefined;
    private _BottomPlaceHolder : PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    console.log("Starting Application Customizer");
    console.log("Available Placeholders: " + this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(", "));

    this._RenderPlaceHolders();

    console.log("Starting Application Customizer");
    return Promise.resolve();
  }
  
  public _RenderPlaceHolders(): void {
    
    if(!this._BottomPlaceHolder){
      this._BottomPlaceHolder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        {onDispose: this._onDispose}
      );
    }

    if(!this._BottomPlaceHolder){
      console.log("Not Able to Fetch {0} Place Holder", PlaceholderName.Bottom);
    }

    let bottomString:string;
    if(this.properties){
      if(!this.properties.Bottom){
        bottomString = "Default Bottom Text being called!";
      }
      else{
        bottomString = this.properties.Bottom;
      }
    }

    if(this._BottomPlaceHolder.domElement){
      this._BottomPlaceHolder.domElement.innerHTML=`<div class="${styles.app}">
      <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.bottom} ${styles.outerStruct}">
        <span style="margin-left:2%; display:inline-flex;">
          <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>Internal Use Only
        </span>
        <span style="margin-right:2%">
          &copy; ${escape(new Date().getFullYear().toString())}
        </span>
      </div>
    </div>`;
    }

  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
