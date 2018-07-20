import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';
import * as strings from 'ValidateWebPartpropertyWebPartStrings';
import ValidateWebPartproperty from './components/ValidateWebPartproperty';
import { IValidateWebPartpropertyProps } from './components/IValidateWebPartpropertyProps';
import { IListInfoWebPartProps } from '../listItems';
import {escape} from '@microsoft/sp-lodash-subset'
import resolveAddress from '../../../temp/workbench-packages/@microsoft_sp-loader/lib/utilities/resolveAddress';

export interface IValidateWebPartpropertyWebPartProps {
  description: string;
}

export default class ValidateWebPartpropertyWebPart extends BaseClientSideWebPart<IListInfoWebPartProps> {

  private validateInput(value:string):Promise<string>{
   return new Promise<string>((resolve:(validationErrorMessage:string)=>void,reject:(error:any)=>void):void=>{
     if(value==null||value.trim().length===0){
       resolve("Field is Empty");
       return;
     }
     this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getByTitle('${escape(value)}')?$select=Id`, SPHttpClient.configurations.v1)
     .then((response:SPHttpClientResponse):void=>{
       if(response.ok){
         resolve('');
         return;
       }else if(response.status===404){
         resolve(`List '${escape(value)}' doesnot exist in the current site.`)
         return;
       }else{
         resolve(`Error: '${escape(response.statusText)}' Please Try Again.`)
         return;
       }
     })
     .catch((error:any):void=>{
       resolve(error);
     });
    });
  }
  public render(): void {
    const element: React.ReactElement<IValidateWebPartpropertyProps > = React.createElement(
      ValidateWebPartproperty,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                  label: strings.DescriptionFieldLabel,
                  onGetErrorMessage:this.validateInput.bind(this)
                }),
                PropertyPaneTextField('listName',{
                  label:strings.ListNameFieldLabel,
                  onGetErrorMessage:this.validateInput.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
