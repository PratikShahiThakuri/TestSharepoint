import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as jQuery from 'jquery';
import styles from './JsEmployeeSpotlightWebPart.module.scss';
import * as strings from 'JsEmployeeSpotlightWebPartStrings';
import {IEmployeeSpotlightWebPartProps} from './IEmployeeSpotlightWebPartProps'
import { SliderHelper } from './Helper';
import { SPHttpClient } from '../../../node_modules/@microsoft/sp-http';

import { PropertyFieldColorPickerMini } from '../../../node_modules/sp-client-custom-fields/lib/PropertyFieldColorPickerMini';
import * as _ from 'lodash';

// export interface IJsEmployeeSpotlightWebPartProps {
//   description: string;
// }

//=================
/**Interface to hold key and value */
export interface ResponceDetails{
  title:string;
  id:string;
}
/**
 * Interface to hold responceDetails Collection.
 */
export interface ResponceDetails{
  value:ResponceDetails[];
}

/**
 *  An interface to hold the ResponceDetails collection.
 */
export interface ResponceCollection {
  value: ResponceDetails[];
}

/**
 * Interface to hold the SpotlightDetails
 */
export interface SpotlightDetails{
  userDisplayName:string;
  userEmail:string;
  userProfilePic:string;
  description:string;
  designation:string;
}
/**
 * class that contain spotlight webpartr operations and corresponding property
 */
export default class JsEmployeeSpotlightWebPart extends BaseClientSideWebPart<IEmployeeSpotlightWebPartProps> {
private spotlightListFirldOptions:IPropertyPaneDropdownOption[]=[];
private spotlightListOptions:IPropertyPaneDropdownOption[]=[];
private siteOptions:IPropertyPaneDropdownOption[]=[];
private defaultProfileImageURL:string='/_layouts/15/userphoto.aspx?size=L';
private helper :SliderHelper=new SliderHelper();
private sliderControl ;any =null;

/**
 * constructor of SpotLightWebPart class
 */
public constructor(){
  super();
  SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.2.1/jquery.min.js', { globalExportsName: 'jQuery' });
//Next button functionality
jQuery(document).on("click","."+styles.next,(event)=>{
  event.preventDefault();
  this.helper.moveSlides(1);
});
//Pervious button functionality
jQuery(document).on("click","."+styles.prev,(event)=>{
  event.preventDefault();
  this.helper.moveSlides(-1);
});
//start stop slider on hover
jQuery(document).ready(()=>{
  jQuery(document).on('mouseenter','.'+styles.containers,()=>{
    if(this.properties.enabledSpotlightAutoPlay){
      clearInterval(this.sliderControl);
    }
  }).on('mouseleave','.'+styles.containers,()=>{
    var carouselSpeed:number =this.properties.spotlightSliderSpeed*1000;
    if(carouselSpeed&& this.properties.enabledSpotlightAutoPlay)
      this.sliderControl=setInterval(this.helper.startAutoPlay,carouselSpeed);
  })
})
}




public render(): void {
    this.domElement.innerHTML = `
      <div id="spListContainer"/>`;
      this._renderSpotlightTemplateAsync();
      this._renderSpotlightDataAsync();
      
      
  }
/**
 * Renders the webpart html with the given spotlight details collection.
 * @param spotLightDetails- a collection of spotlight details.
 */
private _addSpotlightTemplateContent(spotLightDetails:SpotlightDetails[]){
  this.domElement.innerHTML='';
  var innerContent:string='';
  for(let i:number=0;i<spotLightDetails.length;i++){
    innerContent+=`
                <div class="${styles.mySlides}">
                <div style="width:100%;">
                <div style="width:36%; float:left;padding:10%;">
                <img style="border-radius:50%;width:90px;" src="${spotLightDetails[i].userProfilePic}"/>
                </div>
                <div style="width:60%; float:left;text-align:left;">
                <h5 style="margin-bottom:0; text-transform:uppercase;text-align:center">${spotLightDetails[i].userDisplayName}</h5>
                <h6 style="text-align:center;">${spotLightDetails[i].designation}</h6>
                <p>${spotLightDetails[i].description}</p>
                </div>
                
                </div>
                </div>
    `
  }
  this.domElement.innerHTML += `<div class="${styles.containers}" id="slideshow" style="background-color:${this.properties.spotLightBGColor}; cursor:pointer; width:100%; !important;padding: 5px;border-radius: 15px;box-shadow: rgba(0,0,0,0.25) 0 0 20px 0;text-align:center;color:${this.properties.spotlightFontColor};">
  `+ innerContent+`
  <a class="${styles.prev}">&#10094;</a>
  <a class="${styles.next}">&#10095;</a>
  </div> `
  

}

 /**
   * Builds the spotlight details collection with necessary details.
   */
  private _renderSpotlightTemplateAsync(): void {
    if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getSiteCollectionRootWeb().then((response) => {
        this.properties.spotlightSiteCollectionURL = response['Url'];
      });
      if (this.properties.spotlightSiteURL && this.properties.spotlightListName && this.properties.spotlightEmployeeEmailColumn && this.properties.spotlightDescriptionColumn) {
        let spotlightDataCollection: SpotlightDetails[] = [];
        this._getSpotlightListData(this.properties.spotlightSiteURL, this.properties.spotlightListName, this.properties.spotlightEmployeeExpirationDateColumn, this.properties.spotlightEmployeeEmailColumn, this.properties.spotlightDescriptionColumn)
          .then((listDataResponse) => {
            var spotlightListData = listDataResponse.value;
            if (spotlightListData) {
              debugger;
              for (var key in listDataResponse.value) {
                var email = listDataResponse.value[key][this.properties.spotlightEmployeeEmailColumn]["EMail"];
                var id = listDataResponse.value[key]["ID"];
                this._getUserImage(email)
                  .then((response) => {
                    spotlightListData.forEach((item: ResponceDetails) => {
                      let userSpotlightDetails: SpotlightDetails = { userDisplayName: "", userEmail: "", userProfilePic: "", description: "" ,designation:""};
                      if (item[this.properties.spotlightEmployeeEmailColumn]["EMail"] == response["Email"]) {
                        var userName = item[this.properties.spotlightEmployeeEmailColumn];
                        var description = item[this.properties.spotlightDescriptionColumn];
                        var userDescription = "";
                        try {
                          userDescription = (description).text();
                        }
                        catch (err) {
                          userDescription = description;
                        }
                        if (userDescription.length > 140) {
                          var displayFormUrl = this.properties.spotlightSiteURL + '/Lists/' + this.properties.spotlightListName + '/DispForm.aspx?ID=' + id;
                          userDescription = userDescription.substring(0, 140) + `&nbsp; <a href="${displayFormUrl}">ReadMore...</a>`;
                        }
                        var displayName = response["DisplayName"];
                        var designationProperty = _.filter(response["UserProfileProperties"], { Key: "SPS-JobTitle" })[0];
                        var designation = designationProperty["Value"] ? designationProperty["Value"] : "";
                        // uses default image if user image not exist 
                        var profilePicture = response["PictureUrl"] != null && response["PictureUrl"] != undefined ? (<string>response["PictureUrl"]).replace("MThumb", "LThumb") : this.defaultProfileImageURL;
                        // var profilePicture = response["PictureUrl"] != null && response["PictureUrl"] != undefined ? (<string>response["PictureUrl"]) : this.defaultProfileImageUrl;
                        profilePicture = '/_layouts/15/userphoto.aspx?accountname=' + displayName + '&size=M&url=' + profilePicture.split("?")[0];
                        userSpotlightDetails = {
                          userDisplayName: response["DisplayName"],
                          userEmail: response["Email"],
                          userProfilePic: profilePicture,
                          description: userDescription,
                          designation: designation
                        };
                        spotlightDataCollection.push(userSpotlightDetails);
                      }
                    });
                    this._addSpotlightTemplateContent(spotlightDataCollection);
                    if (this.sliderControl == null && this.properties && this.properties.enabledSpotlightAutoPlay) {
                      setTimeout(this.helper.moveSlides(), 2000);
                      this.sliderControl = setInterval(this.helper.startAutoPlay, this.properties.spotlightSliderSpeed * 1000);
                    }
                  });
              }
            }
          });
      }
    }
  }
/**
 * Generic utility function to execute the rest api call and return the curresponding result
 */
private _callAPI(url:string):Promise<ResponceCollection>{
  return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
.then((response)=>{
return response.json();
})  



}
/**
 * Returns  a
 * promise that return user image url and name and email  string containing user email id
 * 
 */
private _getUserImage(email:string){
  return this._callAPI(this.properties.spotlightSiteCollectionURL+"/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|"+email+"'");

}
/**
 * returns a promise that return sitecollection rootweb url.
 */
private _getSiteCollectionRootWeb():Promise<ResponceCollection>{
  return this._callAPI(this.context.pageContext.web.absoluteUrl+`/_api/Site/RootWeb?$select=Title,Url`);
}
/**
 * Returns a Promise that returns all the subsite names from the given site 
 */
private _getAllSubSites(spotlightSiteCollectionURL:string):Promise<ResponceCollection>{
  return this._callAPI(spotlightSiteCollectionURL+`/_api/web/webs?$select=Title,Url`)
}
/**
 * Returns a promise that returns all the list name in corresponding site.
 */
private _getAllList(siteUrl:string):Promise<ResponceCollection>{
  return this._callAPI(siteUrl+`/_api/web/lists?$orderby=Id desc&$filter= Hidden eq false and BaseTemplate eq 100`);
}
/**
   * Returns a promise that returns spotlight list items.
   * @param siteUrl - The string containing site url.
   * @param spotlightListName - The string containing spotlight list name.
   */
  private _getSpotlightListData(siteUrl: string, spotlightListName: string, expiryDateColumn: string, emailColumn: string, descriptionColumn: string): Promise<ResponceCollection> {
    if (siteUrl != "" && spotlightListName != "") {
      var today: Date = new Date();
      var dd: any = today.getDate();
      var mm: any = today.getMonth() + 1; //January is 0!
      var yyyy: any = today.getFullYear();
      dd = (dd < 10) ? '0' + dd : dd;
      mm = (mm < 10) ? '0' + mm : mm;
      var dateString: string = `${yyyy}-${mm}-${dd}`;
      emailColumn = emailColumn.replace(" ", "_x0020_");
      descriptionColumn = descriptionColumn.replace(" ", "_x0020_");
      expiryDateColumn = expiryDateColumn.replace(" ", "_x0020_");
      return this._callAPI(siteUrl + `/_api/web/lists/GetByTitle('${spotlightListName}')/items?$select=ID,${descriptionColumn},${emailColumn}/EMail&$expand=${emailColumn}/Id&$orderby=Id desc&$filter=${expiryDateColumn} ge '${dateString}'`);
    }
  }

/**
 * Loads all the subsites in the current sitecollection and initiates the corresponding drop values loading.
 */
private _renderSpotlightDataAsync():void{
  this._getSiteCollectionRootWeb().then((response)=>{
    this.properties.spotlightSiteCollectionURL = response['Url'];
    this._getAllSubSites(response['Url'])
    .then((siteResponse)=>{
      this.siteOptions=this._getDropDownCollection(siteResponse,'Url','Title');
      this.context.propertyPane.refresh();
      if(this.properties.spotlightSiteURL!=""){
        this._loadAllListsDropDown(this.properties.spotlightSiteURL);
      }
      if(this.properties.spotlightListName!=""){
        this._loadSpotlightListFieldsDropDown(this.properties.spotlightSiteURL,this.properties.spotlightListName)
       
      }
    });
  });
}
/**
 * Load the spotlight list fields for fields dropdown
 */
private _loadSpotlightListFieldsDropDown(siteUrl:string,spotlightListName:string):void{
  this._getSpotlightListFields(siteUrl,spotlightListName)
  .then((response)=>{
    this.spotlightListFirldOptions=this._getDropDownCollection(response,'Title','Title');
    this.context.propertyPane.refresh();
  })
}

/**
 
 * Loads all the list name in selected site for spotlight list dropdown
 */
private _loadAllListsDropDown(siteUrl:string):void{
  this._getAllList(siteUrl)
  .then((response)=>{
    this.spotlightListOptions=this._getDropDownCollection(response,'Title','Title');
    this.context.propertyPane.refresh();
  })
}
/**
 * Returns a promise that returns spotlight field names.
 */
private _getSpotlightListFields(siteUrl:string, spotlightListName:string):Promise<ResponceCollection>{
  if(siteUrl!="" && spotlightListName!="" && siteUrl!=undefined && spotlightListName!=undefined){
    return this._callAPI(siteUrl+`/_api/web/lists/GetByTitle('${spotlightListName}')/Fields?$orderby=Id desc&$filter=Hidden eq false and ReadOnlyField eq false`)
  }
 
}

private _getDropDownCollection(response:ResponceCollection,key:string,text:string){
  var dropdownOptions:IPropertyPaneDropdownOption[]=[];
  if(key=='Url')
    dropdownOptions.push({ key: this.context.pageContext.web.absoluteUrl, text: 'This Site' });
    for(var itemkey in response.value){
      dropdownOptions.push({ key: response.value[itemkey][key], text: response.value[itemkey][text] });

    }
    return dropdownOptions;
  
}
/**
 * validate the property pane fields for null checks
 */
private _validateFiledValue(value:string):string{
  var validationMessage:string='';
  if(value===null||value.trim().length===0){
    validationMessage="please select a value";
  }
  return validationMessage;
}
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart():void{
    if(this.spotlightListFirldOptions.length>0) return;
    this._renderSpotlightTemplateAsync();
     
    
  }
/**
   * Retuns the property pane configuration.
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.propertyPaneHeading,
              groupFields: [
                PropertyPaneDropdown('spotlightSiteURL', {
                  label: strings.selectSiteLableMessage,
                  options: this.siteOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightListName', {
                  label: strings.selectListLableMessage,
                  options: this.spotlightListOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightEmployeeEmailColumn', {
                  label: strings.employeeEmailcolumnLableMessage,
                  options: this.spotlightListFirldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightDescriptionColumn', {
                  label: strings.descriptioncolumnLableMessage,
                  options: this.spotlightListFirldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                }),
                PropertyPaneDropdown('spotlightEmployeeExpirationDateColumn', {
                  label: strings.expirationDateColumnLableMessage,
                  options: this.spotlightListFirldOptions,
                  selectedKey: this._validateFiledValue.bind(this)
                })
              ]
            },
            {
              groupName: strings.effectsGroupName,
              groupFields: [
                PropertyFieldColorPickerMini('spotlightBGColor',{
                  label: strings.spotlightBGColorLableMessage,
                  initialColor: this.properties.spotLightBGColor,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties:this.properties,
                  render:this.render.bind(this.properties),
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'spotlightBGColorFieldId'
                }),
                PropertyFieldColorPickerMini('spotlightFontColor', {
                  label: strings.spotlightFontColorLableMessage,
                  initialColor: this.properties.spotlightFontColor,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this.properties),
                  properties: this.properties,
                  render:this.render.bind(this.properties),
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'spotlightFontColorFieldId'
                }),
                PropertyPaneToggle('enabledSpotlightAutoPlay', {
                  label: strings.enableAutoSlideLableMessage
                }),
                PropertyPaneSlider('spotlightSliderSpeed', {
                  label: strings.carouselSpeedLableMessage,
                  min: 0,
                  max: 7,
                  value: 3,
                  showValue: true,
                  step: 0.5
                })
              ]
            }
          ]
        }
      ]
    };
  }
  /**
   * Triggers foreach Property pane value update and loads the corresponding details.
   * @param propertyPath - The string containing property path.
   * @param oldValue - The string containing old value of property.
   * @param newValue - The string containing new value of property.
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    switch (propertyPath) {
      case "spotlightSiteURL":
        this.properties.spotlightListName = "";
        this._loadAllListsDropDown(this.properties.spotlightSiteURL);
        break;
      case "spotlightListName":
        this.properties.spotlightEmployeeEmailColumn = "";
        this.properties.spotlightDescriptionColumn = "";
        this._loadSpotlightListFieldsDropDown(this.properties.spotlightSiteURL, this.properties.spotlightListName);
        break;
      default:
        break;
    }
  } 
}
