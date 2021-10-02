import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {SPHttpClient,SPHttpClientResponse,ISPHttpClientOptions} from '@microsoft/sp-http'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import jQuery from 'jquery';

import * as strings from 'VenCalendarWebPartStrings';
import VenCalendar from './components/VenCalendar';
import { IVenCalendarProps } from './components/IVenCalendarProps';
import emptyHtml from './components/CalendarTemplate';
import { times } from 'lodash';
import { Fabric, ICalendar } from 'office-ui-fabric-react';
import { ICalendarInfo, ISchedule } from 'tui-calendar';
import * as moment from 'moment';
import * as momentTz from 'moment-timezone';

export interface IVenCalendarWebPartProps {
  site:string;
  siteOther: string;
  listTitle: string;
  other: boolean;
  description: string;
  site2:string;
  siteOther2:string;
  listTitle2:string;
  other2:boolean;
  categoryColumn:string;
  categoryColors:string;
  calendar1Color:string;
  calendar2Color:string;
  filterQuery:string;
}

export interface ISPLists{
  value:ISPList[];
}
export interface ISPList{
  Title:string;
  Id:string;
}

export default class VenCalendarWebPart extends BaseClientSideWebPart <IVenCalendarWebPartProps> {

  public render(): void {
    if (!this.properties.other) {
      jQuery("input[aria-label=hide-col]").parent().hide();
    }
    if (!this.properties.other2) {
      jQuery("input[aria-label=hide-col2]").parent().hide();
    }
    if (this.properties.listTitle == null) {
       var element:React.ReactElement = React.createElement(emptyHtml, {title:this.properties.description});
       ReactDom.render(element, this.domElement);
    } else {
      this._renderListAsync();
    }    
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  //@ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  /*Site 2 variables */
  private colsDisabled: boolean = true;
  private listDisabled: boolean = true;
  private _siteOptions: IPropertyPaneDropdownOption[] = [];
  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];
  private _columnOptions: IPropertyPaneDropdownOption[] = [];
  /*Site 2 variables */
  private colsDisabled2: boolean = true;
  private listDisabled2: boolean = true;
  private _siteOptions2: IPropertyPaneDropdownOption[] = [];
  private _dropdownOptions2: IPropertyPaneDropdownOption[] = [];
  private _columnOptions2: IPropertyPaneDropdownOption[] = [];
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    var otherSiteAria = "hide-col";
    var otherSiteAria2 = "hide-col2";
    if (this.properties.other) {
      otherSiteAria = "";
    }
    if(this.properties.other2){
      otherSiteAria2 = "";
    }
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
                }),
                PropertyPaneDropdown("site", {
                  label: "Site",
                  options: this._siteOptions,
                }),
                PropertyPaneTextField("siteOther", {
                  label:
                    "Other Site Url (i.e. https://contoso.sharepoint.com/path)",
                  ariaLabel: otherSiteAria,
                }),
                PropertyPaneDropdown("listTitle",{
                  label:"List Title",
                  options:this._dropdownOptions,
                  disabled:this.listDisabled
                }),
                PropertyPaneTextField("calendar1Color",{
                  label:"Calendar 1 Color"
                }),
                PropertyPaneDropdown("site2", {
                  label: "Site 2",
                  options: this._siteOptions2,
                }),
                PropertyPaneTextField("siteOther2", {
                  label:
                    "Other Site 2 Url (i.e. https://contoso.sharepoint.com/path)",
                    ariaLabel: otherSiteAria2,
                }),
                PropertyPaneDropdown("listTitle2",{
                  label:"List Title 2",
                  options:this._dropdownOptions2,
                  disabled:this.listDisabled2
                }),
                PropertyPaneTextField("calendar2Color",{
                  label:"Calendar 2 Color"
                }),
                PropertyPaneTextField("categoryColumn",{
                  label:"Category Column"
                }),
                PropertyPaneTextField("categoryColors",{
                  label:"Category Colors",
                  description:"Eg: Business=blue;Meeting=silver"                  
                }),
                PropertyPaneTextField("filterQuery",{
                  label:"Filter Query",
                  description:"Eg: Category eq 'Meeting'"                  
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneConfigurationStart():void{
    if(this.properties.site){
      this.listDisabled = false;
    }
    if (!this.properties.other) {
      jQuery("input[aria-label=hide-col]").parent().hide();
    }
    if(this.properties.site2){
      this.listDisabled2 = false;
    }
    if (!this.properties.other2) {
      jQuery("input[aria-label=hide-col2]").parent().hide();
    }
    if (
      this.properties.site &&
      this.properties.listTitle
    ) {
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Configuration"
      );
      this._getSiteRootWeb().then((response0) => {
        this._getSites(response0["Url"]).then((response) => {
          var sites: IPropertyPaneDropdownOption[] = [];
          sites.push({
            key: this.context.pageContext.web.absoluteUrl,
            text: "This Site",
          });
          sites.push({ key: "other", text: "Other Site (Specify Url)" });
          for (var _key in response.value) {
            if (
              this.context.pageContext.web.absoluteUrl !=
              response.value[_key]["Url"]
            ) {
              sites.push({
                key: response.value[_key]["Url"],
                text: response.value[_key]["Title"],
              });
            }
          }
          this._siteOptions = sites;
          if (this.properties.site) {
            this._getListTitles(this.properties.site).then((response2) => {
              this._dropdownOptions = response2.value.map((list: ISPList) => {
                return {
                  key: list.Title,
                  text: list.Title,
                };
              });
              this.context.propertyPane.refresh();
                this.context.statusRenderer.clearLoadingIndicator(
                  this.domElement
                );
              this.render();              
            });
          }
        });
      });
    } else {
      this._getSitesAsync();
    }
    if (
      this.properties.site2 &&
      this.properties.listTitle2
    ) {
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Configuration"
      );
      this._getSiteRootWeb().then((response0) => {
        this._getSites(response0["Url"]).then((response) => {
          var sites: IPropertyPaneDropdownOption[] = [];
          sites.push({
            key: this.context.pageContext.web.absoluteUrl,
            text: "This Site",
          });
          sites.push({ key: "other2", text: "Other Site (Specify Url)" });
          for (var _key in response.value) {
            if (
              this.context.pageContext.web.absoluteUrl !=
              response.value[_key]["Url"]
            ) {
              sites.push({
                key: response.value[_key]["Url"],
                text: response.value[_key]["Title"],
              });
            }
          }
          this._siteOptions2 = sites;
          if (this.properties.site2) {
            this._getListTitles(this.properties.site2).then((response2) => {
              this._dropdownOptions2 = response2.value.map((list: ISPList) => {
                return {
                  key: list.Title,
                  text: list.Title,
                };
              });
              this.context.propertyPane.refresh();
                this.context.statusRenderer.clearLoadingIndicator(
                  this.domElement
                );
              this.render();              
            });
          }
        });
      });
    } else {
      this._getSites2Async();
    }
  //  this._getSitesAsync();
  //  this._getSites2Async();
  }
  protected onPropertyPaneFieldChanged(
    propertyPath:string,
    oldValue:any,
    newValue:any
  ):void{
    if (newValue == "other") {
      this.properties.other = true;
      this.properties.listTitle = null;
      jQuery("input[aria-label=hide-col]").parent().show();
    } else if (oldValue == "other" && newValue != "other") {
      this.properties.other = false;
      this.properties.siteOther = null;
      this.properties.listTitle = null;
      jQuery("input[aria-label=hide-col]").parent().hide();
    }
    if (newValue == "other2") {
      this.properties.other2 = true;
      this.properties.listTitle2 = null;
      jQuery("input[aria-label=hide-col2]").parent().show();
      //document.querySelector("input[aria-label=hide-col2]").parentElement.hidden = false;
    } else if (oldValue == "other2" && newValue != "other2") {
      this.properties.other2 = false;
      this.properties.siteOther2 = null;
      this.properties.listTitle2 = null;
      jQuery("input[aria-label=hide-col2]").parent().hide();
      //document.querySelector("input[aria-label=hide-col2]").parentElement.hidden = true;
    }
   
    if((propertyPath === "site" || propertyPath === "siteOther")&& newValue){
      this.colsDisabled=true;
      this.listDisabled=true;
      var siteUrl = newValue;
      if(this.properties.other){
        siteUrl = this.properties.siteOther;
      }else{
        jQuery("input[aria-label=hide-col]").parent().hide();
      }
      if((this.properties.other && this.properties.siteOther.length>25 )||(!this.properties.other)){
        this.context.statusRenderer.displayLoadingIndicator(
          this.domElement,
          "Configuration"
        );
        this._getListTitles(siteUrl).then((response)=>{
          this._dropdownOptions = response.value.map((list:ISPList)=>{
            return {
              key: list.Title,
              text: list.Title,
            };
          });
          this.listDisabled=false;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement)
          this.render();
        });
      }
    }else {
      //Handle other fields here
      this.render();
    }
    if((propertyPath === "site2" || propertyPath === "siteOther2")&& newValue){
      this.colsDisabled2=true;
      this.listDisabled2=true;
      var siteUrl = newValue;
      if(this.properties.other2){
        siteUrl = this.properties.siteOther2;
      }else{
        jQuery("input[aria-label=hide-col2]").parent().hide();
        //document.querySelector("input[aria-label=hide-col2]").parentElement.hidden = true;
      }
      if((this.properties.other2 && this.properties.siteOther2.length>25 )||(!this.properties.other2)){
        this.context.statusRenderer.displayLoadingIndicator(
          this.domElement,
          "Configuration"
        );
        this._getListTitles(siteUrl).then((response)=>{
          this._dropdownOptions2 = response.value.map((list:ISPList)=>{
            //if(((!this.properties.other) &&this.properties.site != this.properties.site2 && list.Title !=this.properties.listTitle)||((this.properties.other) &&this.properties.siteOther != this.properties.siteOther2 && list.Title !=this.properties.listTitle)){
              return{
                key:list.Title,
                text:list.Title,
              };
            // }else{
            //   return;
            // }
          });
          this.listDisabled2=false;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement)
          this.render();
        });
      }else {
        //Handle other fields here
        this.render();
      }
    }
  }
  private _getSiteRootWeb(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/Site/RootWeb?$select=Title,Url`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getSites(rootWebUrl: string): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        rootWebUrl + `/_api/web/webs?$select=Title,Url`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  private _getSiteTimeZone(rootWebUrl: string): Promise<any> {
    return this.context.spHttpClient
      .get(
        rootWebUrl + `/_api/Web/RegionalSettings/TimeZone`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  private _getListTitles(site:string):Promise<ISPLists>{
    return this.context.spHttpClient.get(
      site+`/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 106`,
      SPHttpClient.configurations.v1)
      .then((response:SPHttpClientResponse)=>{
        return response.json();
      });
  }
  private _getListData(listName: string, site: string): Promise<any> {   
    var n = this, r = new Headers;
    r.append("odata-version", "3.0");
    var i:ISPHttpClientOptions = {
              headers: r,
              body: "{'query': {\n          '__metadata': {'type': 'SP.CamlQuery'},\n          'ViewXml': '<View><Query><Where><DateRangesOverlap><FieldRef Name=\"EventDate\" /><FieldRef Name=\"EndDate\" /><FieldRef Name=\"RecurrenceID\" /><Value Type=\"DateTime\"><Year/></Value></DateRangesOverlap></Where></Query><QueryOptions><ExpandRecurrence>true</ExpandRecurrence><CalendarDate><Today/></CalendarDate><ViewAttributes Scope=\"RecursiveAll\"/><RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion></QueryOptions></View>'\n        }}"
            }
    let camlQueryPayLoad: any = { query :{__metadata: { type: 'SP.CamlQuery' }, ViewXml: '<View><Query><Where><DateRangesOverlap><FieldRef Name="EventDate" /><FieldRef Name="EndDate" /><FieldRef Name="RecurrenceID" /><Value Type="DateTime"><Year/></Value></DateRangesOverlap></Where></Query><QueryOptions><ExpandRecurrence>true</ExpandRecurrence><ExpandFieldValuesAsText>true</ExpandFieldValuesAsText><CalendarDate><Today/></CalendarDate><ViewAttributes Scope="RecursiveAll"/><RecurrencePatternXMLVersion>v3</RecurrencePatternXMLVersion></QueryOptions></View>'}};

      let spOpts:ISPHttpClientOptions = { 
          headers:{
            "odata-version": "3.0"
          },                 
          body: JSON.stringify(camlQueryPayLoad)  
      };
      var filterQuery = "";
       if( this.properties.filterQuery){filterQuery = "&$filter=" + this.properties.filterQuery;}
    return this.context.spHttpClient
      .post(
        site +
          `/_api/web/lists/GetByTitle('${listName}')/GetItems?$select=Id,Title,EventDate,EndDate,Description,Location,${this.properties.categoryColumn},fAllDayEvent,fRecurrence,RecurrenceData,Duration,FieldValuesAsText/StartDate,FieldValuesAsText/EndDate,Created,Author/ID,Author/Title,XMLTZone,TimeZone&expand=FieldValuesAsText&$limit=500${filterQuery}`,
        SPHttpClient.configurations.v1,spOpts
      )
      .then((response: SPHttpClientResponse) => {
        if(response.ok){
          return response.json();
        }
      }).catch((err) => {
        console.log(err);
      });
  }
  private _getSitesAsync(): void {
    this._getSiteRootWeb().then((response) => {
      this._getSites(response["Url"]).then((response1) => {
                
        var sites: IPropertyPaneDropdownOption[] = [];
        sites.push({
          key: this.context.pageContext.web.absoluteUrl,
          text: "This Site",
        });
        sites.push({ key: "other", text: "Other Site (Specify Url)" });
        for (var _key in response1.value) {
          sites.push({
            key: response1.value[_key]["Url"],
            text: response1.value[_key]["Title"],
          });
        }
        this._siteOptions = sites;
        this.context.propertyPane.refresh();
        var siteUrl = this.properties.site;
        if(this.properties.other){
          siteUrl = this.properties.siteOther;
        }
        this._getListTitles(siteUrl).then((response2)=>{
          this._dropdownOptions =  response2.value.map((list:ISPList)=>{
            return{
                key:list.Title,
                text:list.Title,
              };           
          });
        });
        this.context.propertyPane.refresh();
      });
    });
  }
  private _getSites2Async(): void {
    this._getSiteRootWeb().then((response) => {
      this._getSites(response["Url"]).then((response1) => {
        var sites: IPropertyPaneDropdownOption[] = [];
        sites.push({
          key: this.context.pageContext.web.absoluteUrl,
          text: "This Site",
        });
        sites.push({ key: "other2", text: "Other Site (Specify Url)" });
        for (var _key in response1.value) {
          sites.push({
            key: response1.value[_key]["Url"],
            text: response1.value[_key]["Title"],
          });
        }
        this._siteOptions2 = sites;
        this.context.propertyPane.refresh();
        var siteUrl = this.properties.site2;
        if(this.properties.other2){
          siteUrl = this.properties.siteOther2;
        }
        this._getListTitles(siteUrl).then((response2)=>{
          this._dropdownOptions2 =  response2.value.map((list:ISPList)=>{
            //if(((!this.properties.other) &&this.properties.site != this.properties.site2 && list.Title !=this.properties.listTitle)||((this.properties.other) &&this.properties.siteOther != this.properties.siteOther2 && list.Title !=this.properties.listTitle)){
              return{
                key:list.Title,
                text:list.Title,
              };
           // }
          });
        });
        this.context.propertyPane.refresh();
      });
    });
  }

  private _renderListAsync(){
    var siteUrl = this.properties.site;
    if (this.properties.other) {
      siteUrl = this.properties.siteOther;
    }
    var siteUrl2 = this.properties.site2;
    if (this.properties.other2) {
      siteUrl2 = this.properties.siteOther2;
    }
    
    this._getListData(this.properties.listTitle, siteUrl)
      .then((response) => { 
        this._getSiteTimeZone(siteUrl).then((Site1TimeZoneresponse)=>{
          var Site1TimeZone = Site1TimeZoneresponse.Information.Bias + Site1TimeZoneresponse.Information.StandardBias;
          if(this.properties.listTitle&&this.properties.listTitle2&&(siteUrl+'/'+this.properties.listTitle)!=(siteUrl2+'/'+this.properties.listTitle2)){
            this._getListData(this.properties.listTitle2, siteUrl2)
            .then((response2) => { 
              this._getSiteTimeZone(siteUrl2).then((Site2TimeZoneresponse)=>{
                var Site2TimeZone = Site2TimeZoneresponse.Information.Bias + Site2TimeZoneresponse.Information.StandardBias;
                this._renderList(response.value,siteUrl,response2.value,siteUrl2,Site1TimeZone,Site2TimeZone); 
              });
              //if(response2.ok)                       
            });
          }else{
          this._renderList(response.value,siteUrl,null, null,Site1TimeZone, null);
         } 
        });         
                
      })
      .catch((err) => {
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.context.statusRenderer.renderError(
          this.domElement,
          "There was an error loading your list, please verify the selected list has Calendar Events or choose a new list."
        );
      });
  }

  private _renderList(items: any[],siteUrl:string,items2: any[],siteUrl2:string,site1TimeOffset, site2TimeOffset): void {   
    
    var currentCtx: any = this.context.pageContext.legacyPageContext;  
    var userProfileTimezone = currentCtx.webTimeZoneData.Bias + currentCtx.webTimeZoneData.StandardBias;
    var calendars : ICalendarInfo[] = [];      
    var schedules : ISchedule[] = [];
    var calCategories = {};
    var catColorProp = this.properties.categoryColors;
    if(catColorProp.indexOf(";")>-1){
     var catColors = catColorProp.split(";");
     for(var i=0;i<catColors.length;i++){
       if(catColors[i].indexOf("=")>-1 ){
        var a = catColors[i].split("=");
        calCategories
        calCategories[a[0]] = a[1];
      }
     }
    }
   
      //  this._renderListAsync().then((item)=>{
      //   schedules.push({
      //     id: item["Id"],
      //     calendarId: '0',
      //     title: item["Title"],
      //     category: 'time',
      //     dueDateClass: '',
      //     start: item["EventDate"],
      //     end: item["EndDate"]
      //   });
      // });
     // var items:any[] = this._renderListAsync();     
      if(items){
        calendars.push({
          id: '0',
          name: this.properties.listTitle,
          bgColor: this.properties.calendar1Color,
          borderColor: this.properties.calendar1Color
        });
      var listGuid = this._getListGUID(items[0]);
       var schEvents = [];
        items.map((item:any) => {
          var eventdate =  moment(item["EventDate"]).utcOffset(site1TimeOffset);          
          this.RecurranceEvents(item,null, null).map((event)=>{
          var localStartDateTime = new Date(new Date(event["EventDate"]).getTime()+((site1TimeOffset*60000)-(userProfileTimezone*60000)));
          var localEndDateTime = new Date(new Date(event["EndDate"]).getTime()+((site1TimeOffset*60000)-(userProfileTimezone*60000)));
          event["EventDate"] = localStartDateTime;
          event["EndDate"] = localEndDateTime;
            schEvents.push(event);
          });           
        });
         schEvents.map((schEvent:any) => {
          var startdate= moment(schEvent["EventDate"]).toISOString();
          var enddate =  moment(schEvent["EndDate"]).toISOString();
          var scheduleAnchorAttri = "";
          var itemId=schEvent["Id"];
          if (Environment.type == EnvironmentType.ClassicSharePoint) {
            scheduleAnchorAttri = "javascript:function CalendarItemDialog(){SP.UI.ModalDialog.showModalDialog({url:'" +siteUrl+"/_layouts/15/listform.aspx?PageType=4&ListId="+listGuid+"&ID=" + itemId + "', title:'Details', dialogReturnValueCallback:function(dialogResult){SP.UI.ModalDialog.RefreshPage(dialogResult)}})}CalendarItemDialog();";
            
          } else {
            scheduleAnchorAttri = listGuid ? siteUrl+'/_layouts/15/Event.aspx?ListGuid='+listGuid+'&ItemId='+schEvent["Id"]+'&source='+window.location.href : siteUrl+'/_layouts/15/listform.aspx?PageType=4&ListId='+listGuid+'&ID=' + itemId + '&source=' + +window.location.href;
              
          }
          //var calitems = this.RecurranceEvents(item,null, null);          
          schedules.push({
            id: itemId,
            calendarId: '0',
            title: schEvent["Title"],
            category: schEvent["fAllDayEvent"]==true?'allday':'time',
            dueDateClass: '',
            start: schEvent["EventDate"],
            end: schEvent["EndDate"],
            isAllDay:schEvent["fAllDayEvent"],
            raw:{
              'scheduleAnchorAttri':scheduleAnchorAttri,
              'CalCategory':schEvent[this.properties.categoryColumn],
              'calCategories':calCategories
            }//,
            //recurrenceRule:item["RecurrenceData"]
          });
        });
      }
      if(items2){
        calendars.push({
          id: '1',
          name: this.properties.listTitle2,
          bgColor: this.properties.calendar2Color,
          borderColor: this.properties.calendar2Color
        });
        var listGuid2 = this._getListGUID(items2[0]);
        var schEvents2 = [];
        items2.map((item2:any) => {
          var eventdate =  moment(item2["EventDate"]).utcOffset(site2TimeOffset);
          var utcdate = new Date(new Date(item2["EventDate"]).getTime()+((site2TimeOffset*60000)-(userProfileTimezone*60000)));
          
          this.RecurranceEvents(item2,null, null).map((event2)=>{
          var localStartDateTime = new Date(new Date(event2["EventDate"]).getTime()+((site2TimeOffset*60000)-(userProfileTimezone*60000)));
          var localEndDateTime = new Date(new Date(event2["EndDate"]).getTime()+((site2TimeOffset*60000)-(userProfileTimezone*60000)));
          event2["EventDate"] = localStartDateTime;
          event2["EndDate"] = localEndDateTime;
          schEvents2.push(event2);
          });           
        });
         schEvents2.map((schEvent2:any) => {
          var startdate= moment(schEvent2["EventDate"]).toISOString();
          var enddate = moment(schEvent2["EndDate"]).toISOString();
          var scheduleAnchorAttri = "";
          var itemId2=schEvent2["Id"];
          if (Environment.type == EnvironmentType.ClassicSharePoint) {
            scheduleAnchorAttri = "javascript:function CalendarItemDialog(){SP.UI.ModalDialog.showModalDialog({url:'" +siteUrl2+"/_layouts/15/listform.aspx?PageType=4&ListId="+listGuid2+"&ID=" + itemId2 + "', title:'Details', dialogReturnValueCallback:function(dialogResult){SP.UI.ModalDialog.RefreshPage(dialogResult)}})}CalendarItemDialog();";
            
          } else {
            scheduleAnchorAttri = listGuid2 ? siteUrl2+'/_layouts/15/Event.aspx?ListGuid='+listGuid2+'&ItemId='+schEvent2["Id"]+'&source='+window.location.href : siteUrl2+'/_layouts/15/listform.aspx?PageType=4&ListId='+listGuid2+'&ID=' + itemId2 + '&source=' + +window.location.href;
              
          }
          //var calitems = this.RecurranceEvents(item,null, null);          
          schedules.push({
            id: itemId2,
            calendarId: '1',
            title: schEvent2["Title"],
            category: schEvent2["fAllDayEvent"]==true?'allday':'time',
            dueDateClass: '',
            start: schEvent2["EventDate"],
            end: schEvent2["EndDate"],
            isAllDay:schEvent2["fAllDayEvent"],
            raw:{
              'scheduleAnchorAttri':scheduleAnchorAttri,
              'CalCategory':schEvent2[this.properties.categoryColumn],
              'calCategories':calCategories
            }//,
            //recurrenceRule:item["RecurrenceData"]
          });
        });
      }

      const element: React.ReactElement<IVenCalendarProps> = React.createElement(
        VenCalendar,
        {
          description: this.properties.description,
          calendars:calendars,
          schedules:schedules,
          calCategories :calCategories,     
        }
      );  
      ReactDom.render(element, this.domElement);      
  }
  private RecurranceEvents(schedule, t, n) {
    if (schedule.fRecurrence) {
        t = t || this.parseDate(schedule.EventDate, schedule.fAllDayEvent),
        n = n || this.parseDate(schedule.EndDate, schedule.fAllDayEvent);
        var r = []
          , i = ["su", "mo", "tu", "we", "th", "fr", "sa"]
          , o = ["first", "second", "third", "fourth"]
          , a = 0
          , s = 0;
        if (-1 != schedule.RecurrenceData.indexOf("<repeatInstances>")) {
            var u = schedule.RecurrenceData.substring(schedule.RecurrenceData.indexOf("<repeatInstances>") + 17);
            a = parseInt(u.substring(0, u.indexOf("<")))
        }
        if (-1 != schedule.RecurrenceData.indexOf("<daily ")) {
            var d = schedule.RecurrenceData.substring(schedule.RecurrenceData.indexOf("<daily "));
            d = d.substring(7, d.indexOf("/>") - 1);
            var l = this.formatString(d);
            if (-1 != l.indexOf("dayFrequency"))
                for (var c = parseInt(l[l.indexOf("dayFrequency") + 1]), m = !0, p = this.parseDate(schedule.EventDate, schedule.fAllDayEvent); m; ) {
                    if (s++,
                    new Date(p).getTime() >= t.getTime()) {
                        var h = new Date(p);
                        h.setSeconds(h.getSeconds() + schedule.Duration);
                        var f = this.cloneObj(schedule);
                        f.EventDate = new Date(p),
                        f.EndDate = h,
                        f.fRecurrence = !1,
                        f.Id = schedule.Id,
                        f.ID = f.Id,
                        r.push(f)
                    }
                    p.setDate(p.getDate() + c),
                    (new Date(p) > n || a > 0 && a <= s) && (m = !1)
                }
            else
                -1 != l.indexOf("weekday") && (schedule.RecurrenceData = schedule.RecurrenceData + "<weekly mo='TRUE' tu='TRUE' we='TRUE' th='TRUE' fr='TRUE' weekFrequency='1' />")
        }
        if (-1 != schedule.RecurrenceData.indexOf("<weekly ")) {
            var d = schedule.RecurrenceData.substring(schedule.RecurrenceData.indexOf("<weekly "));
            d = d.substring(8, d.indexOf("/>") - 1);
            for (var l = this.formatString(d), c = parseInt(l[l.indexOf("weekFrequency") + 1]), m = !0, p = this.parseDate(schedule.EventDate, schedule.fAllDayEvent), _ = p.getDay(); m; ) {
                for (var y = _; y < 7; y++)
                    if (-1 != l.indexOf(i[y]) && (a > s || 0 == a) && (s++,
                    new Date(p).getTime() >= t.getTime())) {
                        var g = new Date(p);
                        g.setDate(g.getDate() + (y - _));
                        var h = new Date(g.toString());
                        h.setSeconds(h.getSeconds() + schedule.Duration);
                        var f = this.cloneObj(schedule);
                        f.EventDate = new Date(g.toString()),
                        f.EndDate = h,
                        f.fRecurrence = !1,
                        f.Id = schedule.Id,
                        f.ID = f.Id,
                        r.push(f)
                    }
                p.setDate(p.getDate() + (7 * c - _)),
                _ = 0,
                (new Date(p) > n || a > 0 && a <= s) && (m = !1)
            }
        }
        if (-1 != schedule.RecurrenceData.indexOf("<monthly ")) {
            var d = schedule.RecurrenceData.substring(schedule.RecurrenceData.indexOf("<monthly "));
            d = d.substring(9, d.indexOf("/>") - 1);
            for (var l = this.formatString(d), c = parseInt(l[l.indexOf("monthFrequency") + 1]), m = !0, p = this.parseDate(schedule.EventDate, schedule.fAllDayEvent), v = parseInt(l[l.indexOf("day") + 1]); m; ) {
                if (s++,
                new Date(p).getTime() >= t.getTime()) {
                    var g = new Date(p);
                    if (g.setDate(v),
                    g.getMonth() == p.getMonth()) {
                        var h = new Date(g.toString());
                        h.setSeconds(h.getSeconds() + schedule.Duration);
                        var f = this.cloneObj(schedule);
                        f.EventDate = new Date(g.toString()),
                        f.EndDate = h,
                        f.fRecurrence = !1,
                        f.Id = schedule.Id,
                        f.ID = f.Id,
                        r.push(f)
                    }
                }
                p.setMonth(p.getMonth() + c),
                (new Date(p) > n || a > 0 && a <= s) && (m = !1)
            }
        }
        if (-1 != schedule.RecurrenceData.indexOf("<monthlyByDay ")) {
            var d = schedule.RecurrenceData.substring(schedule.RecurrenceData.indexOf("<monthlyByDay "));
            d = d.substring(14, d.indexOf("/>") - 1);
            for (var l = this.formatString(d), c = parseInt(l[l.indexOf("monthFrequency") + 1]), m = !0, p = this.parseDate(schedule.EventDate, schedule.fAllDayEvent), M = l[l.indexOf("weekdayOfMonth") + 1], b = new Date; m; ) {
                if (s++,
                new Date(p).getTime() >= t.getTime()) {
                    var g = new Date(p);
                    if (g.setDate(1),
                    -1 != l.indexOf("weekday"))
                        if (0 == g.getDay() ? g.setDate(g.getDate() + 1) : 6 == g.getDay() && g.setDate(g.getDate() + 2),
                        "last" == M) {
                            for (; g.getMonth() == p.getMonth(); )
                                b = new Date(g.toString()),
                                5 == g.getDay() ? g.setDate(g.getDate() + 3) : g.setDate(g.getDate() + 1);
                            g = new Date(b.toString())
                        } else
                            for (var L = 0; L < o.indexOf(M); L++)
                                5 == g.getDay() ? g.setDate(g.getDate() + 3) : g.setDate(g.getDate() + 1);
                    else if (-1 != l.indexOf("weekend_day"))
                        if (0 != g.getDay() && 6 != g.getDay() && g.setDate(g.getDate() + (6 - g.getDay())),
                        "last" == M) {
                            for (; g.getMonth() == p.getMonth(); )
                                b = new Date(g.toString()),
                                0 == g.getDay() ? g.setDate(g.getDate() + 6) : g.setDate(g.getDate() + 1);
                            g = new Date(b.toString())
                        } else
                            for (var L = 0; L < o.indexOf(M); L++)
                                0 == g.getDay() ? g.setDate(g.getDate() + 6) : g.setDate(g.getDate() + 1);
                    else if (-1 != l.indexOf("day"))
                        if ("last" == M) {
                            g.setMonth(g.getMonth() + 1);
                            g.setDate(0)
                        } else
                            g.setDate(g.getDate() + o.indexOf(M));
                    else {
                        for (var L = 0; L < i.length; L++)
                            -1 != l.indexOf(i[L]) && (g.getDate() > L ? g.setDate(g.getDate() + (7 - (g.getDay() - L))) : g.setDate(g.getDate() + (y - g.getDay())));
                        if ("last" == M) {
                            for (; g.getMonth() == p.getMonth(); )
                                b = new Date(g.toString()),
                                g.setDate(g.getDate() + 7);
                            g = new Date(b.toString())
                        } else
                            for (var L = 0; L < o.indexOf(M); L++)
                                g.setDate(g.getDate() + 7)
                    }
                    if (g.getMonth() == p.getMonth()) {
                        var h = new Date(g.toString());
                        h.setSeconds(h.getSeconds() + schedule.Duration);
                        var f = this.cloneObj(schedule);
                        f.EventDate = new Date(g.toString()),
                        f.EndDate = h,
                        f.fRecurrence = !1,
                        f.Id = schedule.Id,
                        f.ID = f.Id,
                        r.push(f)
                    }
                }
                p.setMonth(p.getMonth() + c),
                (new Date(p) > n || a > 0 && a <= s) && (m = !1)
            }
        }
        if (-1 != schedule.RecurrenceData.indexOf("<yearly ")) {
            var d = schedule.RecurrenceData.substring(schedule.RecurrenceData.indexOf("<yearly "));
            d = d.substring(8, d.indexOf("/>") - 1);
            for (var l = this.formatString(d), c = parseInt(l[l.indexOf("yearFrequency") + 1]), m = !0, p = this.parseDate(schedule.EventDate, schedule.fAllDayEvent), D = parseInt(l[l.indexOf("month") + 1]) - 1, v = parseInt(l[l.indexOf("day") + 1]); m; ) {
                var g = new Date(p);
                if (g.setMonth(D),
                g.setDate(v),
                new Date(p).getTime() <= g.getTime() && (s++,
                new Date(p).getTime() >= t.getTime())) {
                    var h = new Date(g.toString());
                    h.setSeconds(h.getSeconds() + schedule.Duration);
                    var f = this.cloneObj(schedule);
                    f.EventDate = new Date(g.toString()),
                    f.EndDate = h,
                    f.fRecurrence = !1,
                    f.Id = schedule.Id,
                    f.ID = f.Id,
                    r.push(f)
                }
                p.setFullYear(p.getFullYear() + c),
                (new Date(p) > n || a > 0 && a <= s) && (m = !1)
            }
        }
        if (-1 != schedule.RecurrenceData.indexOf("<yearlyByDay ")) {
            var d = schedule.RecurrenceData.substring(schedule.RecurrenceData.indexOf("<yearlyByDay "));
            d = d.substring(13, d.indexOf("/>") - 1);
            for (var l = this.formatString(d), c = parseInt(l[l.indexOf("yearFrequency") + 1]), m = !0, p = this.parseDate(schedule.EventDate, schedule.fAllDayEvent), D = parseInt(l[l.indexOf("month") + 1]) - 1, M = l[l.indexOf("weekdayOfMonth") + 1], v = 0, L = 0; L < i.length; L++)
                -1 != l.indexOf(i[L]) && "true" == l[l.indexOf(i[L]) + 1].toLowerCase() && (v = L);
            for (; m; ) {
                var g = new Date(p);
                if (g.setMonth(D),
                new Date(p).getTime() <= g.getTime() && (s++,
                new Date(p).getTime() >= t.getTime())) {
                    g.setDate(1);
                    var w = g.getDay();
                    if (v < w ? g.setDate(g.getDate() + (7 - w + v)) : g.setDate(g.getDate() + (v - w)),
                    "last" == M)
                        for (var b = new Date(g.toString()); b.getMonth() == D; )
                            g = new Date(b.toString()),
                            b.setDate(b.getDate() + 7);
                    else
                        g.setDate(g.getDate() + 7 * o.indexOf(M));
                    if (g.getMonth() == D) {
                        var h = new Date(g.toString());
                        h.setSeconds(h.getSeconds() + schedule.Duration);
                        var f = this.cloneObj(schedule);
                        f.EventDate = new Date(g.toString()),
                        f.EndDate = h,
                        f.fRecurrence = !1,
                        f.Id = schedule.Id,
                        f.ID = f.Id,
                        r.push(f)
                    }
                }
                p.setFullYear(p.getFullYear() + c),
                p.setMonth(D),
                p.setDate(1),
                (new Date(p) > n || a > 0 && a <= s) && (m = !1)
            }
        }
        return r
    }
    return schedule.EventDate = new Date(this.parseDate(schedule.EventDate, schedule.fAllDayEvent)),
    schedule.EndDate = new Date(this.parseDate(schedule.EndDate, schedule.fAllDayEvent)),
    [schedule]
}
private parseDate(date, isAllDayEvent) {
  return "string" == typeof date ? new Date(date) : date;
}
private formatString(e) {
  var t = e.split("'");
  return e = t.join(""),
  t = e.split('"'),
  e = t.join(""),
  t = e.split("="),
  e = t.join(" "),
  e.trim(),
  e.split(" ")
}
private cloneObj(e) {
  var t;
  if (null == e || "object" != typeof e)
      return e;
  if (e instanceof Date)
      return t = new Date,
      t.setTime(e.getTime()),
      t;
  if (e instanceof Array) {
      t = [];
      for (var n = 0, r = e.length; n < r; n++)
          t[n] = this.cloneObj(e[n]);
      return t
  }
  if (e instanceof Object) {
      t = {};
      for (var i in e)
          e.hasOwnProperty(i) && (t[i] = this.cloneObj(e[i]));
      return t
  }
  throw new Error("Unable to copy obj! Its type isn't supported.")
}
 private _getListGUID(listItem) {
    var listGuid = "";
    if (console.log(listItem),listItem) {
        var odataEditLink = listItem["odata.editLink"]
          , startGuidChar = odataEditLink.indexOf("'")
          , lastGuidChar = odataEditLink.indexOf("'", startGuidChar + 1);
          odataEditLink = odataEditLink.substring(startGuidChar + 1, lastGuidChar),
        listGuid = odataEditLink
    } //else
    //     t = !1;
    return listGuid
  }
}

