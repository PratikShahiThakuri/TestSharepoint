
import { resolvePath } from "../../temp/workbench-packages/@microsoft_sp-loader/lib/utilities/resolveAddress";
import { IDocument } from "./IDocument";

export class DocumentsService{
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
    public static getRecentDocument():Promise<IDocument>{
        return new Promise<IDocument>((resolve:(document:IDocument)=>void,reject:(error:any)=>void):void=>{
            window.setTimeout(():void => {
                resolve(DocumentsService.documents[0]);
            }, 300);
        })
    }
    public static getRecentDocuments(startFrom:number=0):Promise<IDocument[]>{
        return new Promise<IDocument[]>((resolve:(documents:IDocument[])=>void,reject:(error:any)=>void)=>{
            window.setTimeout(():void=>{
                resolve(DocumentsService.documents.slice(startFrom,startFrom+3));
            },300)
        })
    }
    private static ensureRecentDocuments():Promise<IDocument[]>{
        return new Promise<IDocument[]>((resolve:(documents:IDocument[])=>void,reject:(error:any)=>void):void=>{
            if((window as any).loadedData){
                resolve((window as any).loadedData)
                return;
            }
            if((window as any).loadingData){
                window.setTimeout(():void=>{
                    DocumentsService.ensureRecentDocuments()
                    .then((recentDocuments:IDocument[]):void=>{
                        resolve(recentDocuments)
                    })
                },100)
            }else{
                (window as any).loadingData=true;
                window.setTimeout(():void=>{
                    (window as any).loadedData = DocumentsService.documents;
                (window as any).loadingData=false;
                resolve((window as any).loadedData);
                },300);

                
            }
        })
    }
}