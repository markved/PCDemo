import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PcdemoWebPart.module.scss';
import * as strings from 'PcdemoWebPartStrings';

export interface IPcdemoWebPartProps {
  description: string;
}
import { IListItem } from './IListItem';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { sp, Item, ItemAddResult, ItemUpdateResult } from '@pnp/sp';

export default class PcdemoWebPart extends BaseClientSideWebPart<IPcdemoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.pcdemo}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to Project Validate Demo!</span>
              <p class="${ styles.subTitle}">Find out if this project is ready to be promoted to next stage</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <p><button class="${styles.button} read-Button">
              <span class="${styles.label}">Validate</span>
            </button></p>
            </div>
          </div>
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <div class="status"></div>
          <ul class="items"><ul>
        </div><div class="moveToActionButtons"></div>
      </div>
      
        </div>
      </div>`;
    this.setButtonsEventHandlers();
  }
  private setButtonsEventHandlers(): void {
    const webPart: PcdemoWebPart = this;
    this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.readItem(); event.preventDefault(); });
    //this.readItem();
  }

  private readItem(): void {
    this.updateStatus('Loading latest items...');
    this.getContextItemId()
      .then((itemId: number): Promise<IListItem> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        this.updateStatus(`Loading information about item ID: ${itemId}...`);
        return sp.web.lists.getByTitle("Projects")
          .items.getById(itemId).select('Title', 'Id', 'PRJworkflowStage', 'l_MSPWAPROJUID', 'PRJowner/EMail').expand('PRJowner').get();
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item ID: ${item.Id}, Title: ${item.Title}, Workflow Stage: ${item.PRJworkflowStage}, Project ID: ${item.l_MSPWAPROJUID}, Owner: ${item["PRJowner"].EMail}, Current User: ${this.context.pageContext.user.email}`);
        if (item["PRJowner"].EMail === this.context.pageContext.user.email){
        this.domElement.querySelector('.moveToActionButtons').innerHTML = `  
              <div class="${ styles.column}">
              <p>Project can be moved to next stage. Click on one the buttons below!</p>
                <p><button class="${styles.button} planning-Button">
                <span class="${styles.label}">Move to Planning</span>
              </button>
              <p><button class="${styles.button} cancel-Button">
                <span class="${styles.label}">Cancel Project</span>
              </button></p> 
              <p><button class="${styles.button} hold-Button">
                <span class="${styles.label}">Put Project on Hold</span>
              </button></p>
              </p>
        </div>`;
        }
        else{
          this.domElement.querySelector('.moveToActionButtons').innerHTML = `  
          <div class="${ styles.column}">
          <p>Project can ONLY be moved to next stage by Project Owner.</p>  
          </div>`;
        }
      }, (error: any): void => {
        this.updateStatus('Loading latest item failed with error: ' + error);
      });
  }
  private updateStatus(status: string, items: IListItem[] = []): void {
    this.domElement.querySelector('.status').innerHTML = status;
    this.updateItemsHtml(items);
  }
  private updateItemsHtml(items: IListItem[]): void {
    this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id})</li>`).join("");
  }
  private getContextItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      var queryParameters = new UrlQueryParameterCollection(window.location.href);
      var projUID = queryParameters.getValue("ProjUid");
      sp.web.lists.getByTitle("Projects")
        .items.filter("l_MSPWAPROJUID eq '" + projUID + "'").orderBy('Id', false).top(1).select('Id').get()
        .then((items: { Id: number }[]): void => {
          if (items.length === 0) {
            resolve(-1);
          }
          else {
            resolve(items[0].Id);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      sp.web.lists.getByTitle("Projects")
        .items.orderBy('Id', false).top(1).select('Id').get()
        .then((items: { Id: number }[]): void => {
          if (items.length === 0) {
            resolve(-1);
          }
          else {
            resolve(items[0].Id);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
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
