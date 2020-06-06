import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import { sp } from "@pnp/sp";
import { Session } from '@pnp/sp-taxonomy';
import { SPHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'ExtensionDemoApplicationCustomizerStrings';
import styles from './AppCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

const LOG_SOURCE: string = 'ExtensionDemoApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IExtensionDemoApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Top: string;
  Bottom: string;
}

interface IFindTermSetRequest {
  searchTerms: string;
  lcid: number;
}

interface ITermSet {
  Id: string;
  Name: string;
  Owner: string;
}

interface ITerm {
  Id: string;
  Label: string;
  Paths: string[];
}

interface IFindTermSetResult {
  Error: string;
  Lm: number;
  Content: any[];
}

interface IGetChildTermsInTermWithPagingRequest {
  sspId: string; // guid of term store
  lcid: number;
  guid: string; // guid of term
  termsetId: string; // guid of term set
  includeDeprecated: boolean;
  pageLimit: number;
  pagingForward: boolean;
  includeCurrentChild: boolean;
  currentChildId: string;
  webId: string;
  listId: string;
}
/** A Custom Action which can be run during execution of a Client Side Application */
export default class ExtensionDemoApplicationCustomizer
  extends BaseApplicationCustomizer<IExtensionDemoApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolder);

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
      console.log("OnInit ran.");
    });
    //return Promise.resolve();
  }

  private _renderPlaceHolder(): void {
    console.log("term data1");
    this.getChildTermsInTermWithPaging();
      
    
  }
  private _onDispose(): void {
    console.log('[ExtensionDemoApplicationCustomizer._onDispose]');
  }
  
  private getChildTermsInTermWithPaging() {
    const url = "https://ajindeveloper.sharepoint.com/"+ '/_vti_bin/TaxonomyInternalService.json/GetChildTermsInTermWithPaging';
    const query: IGetChildTermsInTermWithPagingRequest = {
      lcid: 1033,
      sspId: "75aa5f8d37e54e06bfda215863dfb824", //id of termstore
      guid: "75aa5f8d37e54e06bfda215863dfb824", //id of term
      termsetId: "de41fdc4-56d4-4eb9-ab43-fc45df193718", //id of termset
      includeDeprecated: false,
      pageLimit: 1000,
      pagingForward: false,
      includeCurrentChild: true,
      currentChildId: "00000000-0000-0000-0000-000000000000",
      webId: "00000000-0000-0000-0000-000000000000",
      listId: "00000000-0000-0000-0000-000000000000"
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          let returnResults: ITerm[] = [];
          console.log( result.d);
        });
      } else {
        console.log(response.statusText);
      }
    });
  }

}
