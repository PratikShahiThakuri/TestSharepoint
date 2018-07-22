import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'RecentDocumentsWebPartStrings';
import RecentDocuments from './components/RecentDocuments';
import { IRecentDocumentsProps } from './components/IRecentDocumentsProps';
import { IDocument } from '../../documentsServices/IDocument';
import { IRecentDocumentProps } from './IRecentDocumentProps';
import { DocumentsService } from '../../documentsServices';
export interface IRecentDocumentsWebPartProps {
  description: string;
}

export default class RecentDocumentsWebPart extends BaseClientSideWebPart<IRecentDocumentsProps> {
  private static documents: IDocument[] = [
    {
        title: 'Proposal for Jacksonville Expansion Ad Campaign',
        url: 'https://www.google.com',
        imageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSFBD4er1rTDH_-bC7XHWZ3U3wgNDiohyy8EwHPq28tVgaGdJXDXg',
        iconUrl: '',
        activity: {
            title: 'Modified, July 22 2018',
            actorName: 'Miriam Graham',
            actorImageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS0_2nqm0H20gpO-Pf9BsBwuAYt3McWcb-6rFs37i244h71Lyrnkg'
        }
    },
    {
        title: 'Customer Feedback for ZT1000',
        url: 'https://www.google.com',
        imageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSFBD4er1rTDH_-bC7XHWZ3U3wgNDiohyy8EwHPq28tVgaGdJXDXg',
         iconUrl: '',
        activity: {
            title: 'Modified, January 23 2017',
            actorName: 'Miriam Graham',
            actorImageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS0_2nqm0H20gpO-Pf9BsBwuAYt3McWcb-6rFs37i244h71Lyrnkg'
            
        }
    },
    {
        title: 'Asia Q3 Marketing Overview',
        url: 'https://www.google.com',
        imageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSFBD4er1rTDH_-bC7XHWZ3U3wgNDiohyy8EwHPq28tVgaGdJXDXg',
       
        iconUrl: '',
        activity: {
            title: 'Modified, January 23 2017',
            actorName: 'Alex Wilber',
            actorImageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS0_2nqm0H20gpO-Pf9BsBwuAYt3McWcb-6rFs37i244h71Lyrnkg'
        }
    },
    {
        title: 'Trey Research Business Development Plan',
        url: 'https://www.google.com',
        imageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSFBD4er1rTDH_-bC7XHWZ3U3wgNDiohyy8EwHPq28tVgaGdJXDXg',
        iconUrl: '',
        activity: {
            title: 'Modified, January 15 2017',
            actorName: 'Alex Wilber',
            actorImageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS0_2nqm0H20gpO-Pf9BsBwuAYt3McWcb-6rFs37i244h71Lyrnkg'
        }
    },
    {
        title: 'XT1000 Marketing Analysis',
        url: 'https://www.google.com',
        imageUrl: 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSFBD4er1rTDH_-bC7XHWZ3U3wgNDiohyy8EwHPq28tVgaGdJXDXg',
        iconUrl: '',
        activity: {
            title: 'Modified, December 15 2016',
            actorName: 'Henrietta Mueller',
            actorImageUrl: 'https://contoso-my.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=henriettam@contoso.onmicrosoft.com&size=L'
        }
    }
];



public render(): void {
  this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'documents');

  DocumentsService.getRecentDocuments().then((documents:IDocument[])=>{
    const elem :React.ReactElement<IRecentDocumentProps>=React.createElement(
      RecentDocuments,
      {
        documents:documents
      }
    );
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      ReactDom.render(elem, this.domElement);
    
    });

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
               
              ]
            }
          ]
        }
      ]
    };
  }
}
