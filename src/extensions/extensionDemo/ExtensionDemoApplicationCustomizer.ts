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
const HEADER_TEXT: string = "This is the top zone";
const FOOTER_TEXT: string = "This is the bottom zone";
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
interface IExtensionDemoApplicationCustomizerProperties {
  // This is an example; replace with your own property
  TopContent: string;
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
      //console.log("OnInit ran.");
    });
    //return Promise.resolve();
  }

  

  private _renderPlaceHolder(): void {
    console.log("term data124422");
    // usage:
    this.getTermsetWithChildren(
      'https://ajindeveloper.sharepoint.com/',
      'Taxonomy_SkzIIXWk3+at2Pc/WQGciA==',
      'de41fdc4-56d4-4eb9-ab43-fc45df193718'
    ).then(data => {
      console.log("got term data");
      console.log(data);
      // top placeholder..
      let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
      if (topPlaceholder) {
      let nav='<ul>';
        (data as any[]).forEach((ele)=>{
          nav+='<li><a href="#'+ele.Name+'">'+ele.Name+'</a></li>';
        });
        nav+="</ul>";
        topPlaceholder.domElement.innerHTML = `<div class="${styles.app}">
                  <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
                    `+nav+`
                  </div>
                </div>`;
      }

      // bottom placeholder..
      let bottomPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
      if (bottomPlaceholder) {
        bottomPlaceholder.domElement.innerHTML = `<div class="${styles.app}">
                  <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.bottom}">
                    <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>&nbsp; ${escape(FOOTER_TEXT)}
                  </div>
                </div>`;
      }
    });


  }
  private _onDispose(): void {
    console.log('[ExtensionDemoApplicationCustomizer._onDispose]');
  }

  private getTermsetWithChildren(siteCollectionURL: string, termStoreName: string, termsetId: string) {
    console.log("term data2");
    return new Promise((resolve, reject) => {
      const taxonomy = new Session(siteCollectionURL);
      const store: any = taxonomy.termStores.getByName(termStoreName);
      console.log(store);      
      store.getTermSetById(termsetId).terms.select('Name', 'Id', 'Parent').get()
        .then((data: any[]) => {
          console.log("got term data");
          let result = [];
          // build termset levels
          do {
            for (let index = 0; index < data.length; index++) {
              let currTerm = data[index];
              if (currTerm.Parent) {
                let parentGuid = currTerm.Parent.Id;
                insertChildInParent(result, parentGuid, currTerm, index);
                index = index - 1;
              } else {
                data.splice(index, 1);
                index = index - 1;
                result.push(currTerm);
              }
            }
          } while (data.length !== 0);
          // recursive insert term in parent and delete it from start data array with index
          function insertChildInParent(searchArray, parentGuid, currTerm, orgIndex) {
            searchArray.forEach(parentItem => {
              if (parentItem.Id == parentGuid) {
                if (parentItem.children) {
                  parentItem.children.push(currTerm);
                } else {
                  parentItem.children = [];
                  parentItem.children.push(currTerm);
                }
                data.splice(orgIndex, 1);
              } else if (parentItem.children) {
                // recursive is recursive is recursive
                insertChildInParent(parentItem.children, parentGuid, currTerm, orgIndex);
              }
            });
          }
          resolve(result);
        }).catch(fail => {
          console.warn(fail);
          reject(fail);
        });
    });
  }


}
