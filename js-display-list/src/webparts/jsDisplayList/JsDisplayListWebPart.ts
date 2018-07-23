import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './JsDisplayListWebPart.module.scss';
import * as strings from 'JsDisplayListWebPartStrings';

import {  Environment, EnvironmentType, Log } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {IJsDisplayListWebPartProps} from './IJsDisplayListWebPartProps';

// export interface IJsDisplayListWebPartProps {
//   description: string;
// }
//===================
export interface ISPLists{
  value:ISPList[];
}
export interface ISPList{
  Title:string;
  Id:string;
  File:File;
}
export interface File{
  Name:string;
  Length:number;
  Created:string;
  Modified:string;
  ServerRelativeUrl:string;

}
export interface ISPOption{
  Id:string;
  Title:string;
}
export default class JsDisplayListWebPart extends BaseClientSideWebPart<IJsDisplayListWebPartProps> {

  public render(): void {
   this.context.statusRenderer.clearError(this.domElement);
   Log.verbose('js-display-list','Invoking render');
   this._renderListAsync();
  }
  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                PropertyPaneDropdown('listTitle', {
                  label: 'List Title',
                  options:this._dropdownOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];
  public onInit<T>():Promise<T>{
    this._getListTitles()
    .then((response)=>{
      this._dropdownOptions=response.value.map((list:ISPList)=>{
        return{
          key:list.Title,
          text:list.Title
        };
      });
    });
    return Promise.resolve();
  }
  private _getListTitles(): Promise<ISPLists> {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
          return response.json();
      });
  }
  private _getListData(listName:string):Promise<ISPLists>{
    const queryString:string ='$select=Title,ID,Created,Modified,Author/ID,Author/Title,File&$expand=Author/ID,Author/Title,File';
    return this.context.spHttpClient
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items?${queryString}`,
    SPHttpClient.configurations.v1
  ).then((response:SPHttpClientResponse)=>{
    if(response.status===404){
      Log.error('js-display-List',new Error('List Not Found.'));
      return[]
    }else{
      return response.json();
    }
  })
  }
  private _renderList(items:ISPList[]):void{
   
    let html:string='';
    if(!items){
      html='<br/><p class="ms-font-m-plus">The selected list doesnot exist.</p>';
    }else if(items.length===0){
      html='<br/><p class="ms-font-m-plus"> The selected list is empty</p>';
    }else{
      items.forEach((item:ISPList)=>{
        console.log(item);
        let title :string='';
        let size :string='';
        let sizeHtml:string='';
        let link : string='';
        let linkhtml:string='';
        if(item.Title===null){
         if(item.File===null && item.Title===null){
          title="Missing title for item with ID= "+ item.Id;
        }else {
          title=item.File.Name;
          size = (item.File.Length/1024).toFixed(2);
          link = item.File.ServerRelativeUrl;
          sizeHtml=`<div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
          ${size}
        </div>`
        linkhtml =`<div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
        <a href="${link}">Link</a>
      </div>`
        }
        }
        else{
          title=item.Title;
        }
        let created:any =item["Created"];
        let modified :any =item["Modified"]
        html+=`
        <div class ="${styles.row} ms-Grid-row ">
        <div class=" ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
        ${title}               
        </div>
        <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
        ${created.substring(0,created.length -1).replace('T',' ')}
        </div>
        <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
        ${modified.substring(0,created.length -1).replace('T',' ')}
        </div>
        <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-font-m">
        ${item['Author'].Title}
        </div>
       ${sizeHtml}
        ${linkhtml}
        
        </div>
        
        
        `
      });
    }
    const listContainer:Element=this.domElement.querySelector("#spListContainer");
    listContainer.innerHTML = html;
  }
  private _renderListAsync(): void {

    this.domElement.innerHTML = `
        <div className='wrapper'>
          <p class="ms-font-l ms-bgColor-themeDark ms-fontColor-white">
          <span class="ms-fontWeight-semibold">
              ${this.properties.listTitle}
              </span>
              List
          </p>
          <div class="ms-Grid ${styles.jsDisplayList}">
             <div class="ms-Grid-row">
                <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-bgColor-themeLight  ms-font-m-plus">Title</div>
                <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-bgColor-themeLight  ms-font-m-plus">Created</div>
                <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2 ms-bgColor-themeLight  ms-font-m-plus">Modified</div>
                <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2  ms-bgColor-themeLight  ms-font-m-plus">Created By</div>
                <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2  ms-bgColor-themeLight  ms-font-m-plus">Size</div>
                <div class="ms-Grid-col ms-u-sm5 ms-u-md3 ms-u-lg2  ms-bgColor-themeLight  ms-font-m-plus">Url</div>

              </div>
              <hr />
              <div id="spListContainer"></div>
        </div>`;

    const listContainer: Element = this.domElement.querySelector('#spListContainer');

    // Local environment
    // debugger;
    if (Environment.type === EnvironmentType.Local) {
      let html: string = '<p> Local test environment [No connection to SharePoint]</p>';
      listContainer.innerHTML = html;
    } else {
      //debugger;
      
      this._getListData(this.properties.listTitle).then((response) => {
      
        this._renderList(response.value);

      }).catch((err) => {
        Log.error('js-display-List', err);
        this.context.statusRenderer.renderError(this.domElement, err);
      });
    }
  }
}
