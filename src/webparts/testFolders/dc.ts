/*
import { 
    Version,
    Log
  //  DisplayMode,
  //  Environment,
  //  EnvironmentType
  
  } from "@microsoft/sp-core-library";
  import {
    IPropertyPaneConfiguration,
    //PropertyPaneTextField,
    //PropertyPaneCheckbox,
    PropertyPaneChoiceGroup,
    //PropertyPaneLabel,
    //PropertyPaneLink,
    //PropertyPaneSlider,
    //PropertyPaneToggle,
    //PropertyPaneButton,
    //PropertyPaneButtonType,
    //PropertyPaneDropdown
  } from "@microsoft/sp-property-pane";
  import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
  import { IReadonlyTheme } from "@microsoft/sp-component-base";
  import { escape } from "@microsoft/sp-lodash-subset";
  import styles from "./DocumentCentreWebPart.module.scss";
  import * as strings from "DocumentCentreWebPartStrings";
  
  import {
    SPHttpClient,
    ISPHttpClientOptions,
    SPHttpClientResponse,
  } from "@microsoft/sp-http";
  
  import { SPComponentLoader } from "@microsoft/sp-loader";
  import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';
  import { MSGraphClientV3 } from "@microsoft/sp-http";
  //import {getSP} from './pnpjsConfig';
  //import { sp } from "@pnp/sp";
  import { Web } from "@pnp/sp/webs";
  import "@pnp/sp/lists";
  import "@pnp/sp/items";
  
  import { IDocumentCentreProps } from "./DocumentCentreProps";
  //import { setLibraryListeners } from './components/Listeners';
  //import {  } from './components/DataFunctions';
  //import { getLibraryTabs } from './components/LibraryTabs';
  //import { checkPowerUserPermission } from './components/Permissions';
  
  // *** import external frameworks / libraries
  import * as $ from 'jquery';
  require("bootstrap");
  require("./styles/custom.css");
  
  // *** set global variables
  let alertContainer: Element;
  let alertHTML: string = "";
  let graphClient: MSGraphClientV3;
  
  export interface ISPLists {
    value: ISPList[];
  }
  
  export interface ISPList {
    Id : string;
    //Team : string;
    Title : string;
  }
  
  export default class DocumentCentreWebPart extends BaseClientSideWebPart<IDocumentCentreProps> {
    //private _isDarkTheme: boolean = false;
    //private _environmentMessage: string = '';
    //private _sp: SPFI;
  
    // *** set Document Centre web URLs for each Division
    private asmDC = Web("https://maximusunitedkingdom.sharepoint.com/sites/asm_dc/");
    private cenDC = Web("https://maximusunitedkingdom.sharepoint.com/sites/cen_dc/");
    private empDC = Web("https://maximusunitedkingdom.sharepoint.com/sites/emp_dc/");
    private heaDC = Web("https://maximusunitedkingdom.sharepoint.com/sites/hea_dc/");
    private cnnDC = Web("https://maximusunitedkingdom.sharepoint.com/sites/cnn_dc/");
  
    // **** Function  : getLibraryTabs
    // **** Purpose   : display the standard document centre library tabs.
    // ****
    // ****
    private async getLibraryTabs(): Promise<void> {
      const library = ["Policies", "Procedures", "Guides", "Forms", "General"];
      let html: string = "";
      //let libraryFlag : boolean=false;
  
      try {
        // *** check if user is in Managers Security Group
        await this.checkManagerPermission();
        if (this.properties.isManager) {
          library.push("Management");
        }
  
        for (let x = 0; x < library.length; x++) {
          //console.log("library URL=" + this.properties.dcURL + "/" + library[x]);
          this.checkDataAsync(library[x],this.properties.siteName,"");
          //console.log("dataFlag="+this.properties.dataFlag);
  
          if(this.properties.dataFlag){
            if (this.properties.isDCPowerUser) {
              html += `<div class="row"> 
                        <button class="btn libraryBtn nav-link text-left mb-1" id="${library[x]}Tab" data-bs-toggle="pill" data-bs-target="#${library[x]}" aria-controls="${library[x]}" aria-selected="true" type="button" role="tab">
                            <div class="col-1 libraryUploadIcon">
                            <a href="${this.properties.dcURL}/${library[x]}/forms/${this.properties.siteShortName}.aspx" target="_blank">
                                <h3 class="text-white"><i class="bi bi-cloud-arrow-up"></i></h3>
                            </a>
                            </div>
                            <div class="col-11 libraryName">
                            <h6 class="libraryText">${library[x]}</h6>
                            </div>
                        </button>
                        </div>`;
            } else {
              html += `<div class="row"><button class="btn libraryBtn nav-link text-left mb-1" id="${library[x]}Tab" data-bs-toggle="pill" data-bs-target="#${library[x]}" type="button" role="tab" aria-controls="${library[x]}" aria-selected="true">${library[x]}</button></div>`;
            }
          }
        }
  
        const listContainer: Element = this.domElement.querySelector("#libraryTabs");
        listContainer.innerHTML = html;
  
        // *** get custom tabs from termstore and add library column
        await this.renderCustomTabsAsync();
      } catch (err) {
        await this.addError(this.properties.siteName,"getLibraryTabs",err.message);
      }
      return;
    }
  
    // **** Function  : checkManagerPermission
    // **** Purpose   : is the current user a manager? check the 365 managers group.
    // ****
    // ****
    private async checkManagerPermission(): Promise<void> {
      let mgrFlag = false;
      try {
        const alertHTML: string ='<div class="alert alert-info" role="alert">Checking for Manager Permission - Please wait...</div>';
        const alertContainer: Element =this.domElement.querySelector("#headerBar");
        alertContainer.innerHTML = alertHTML;
  
        this.properties.isManager = false;
        this.properties.managersGroupID = "c3da59d5-cb8b-4df2-94ab-3213e5e3a1b0";
  
        //const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient("3");
        graphClient = await this.context.msGraphClientFactory.getClient("3");
        const members = await graphClient.api("/groups/c3da59d5-cb8b-4df2-94ab-3213e5e3a1b0/members").top(999).get();
        const myDetails = await graphClient.api("/me").get();
  
        for (let x = 0; x < members.value.length; x++) {
          if (members.value[x].id === myDetails.id) {
            this.properties.isManager = true;
  
            mgrFlag = true;
            break;
          }
        }
        alertContainer.innerHTML = "";
        this.properties.isManager = mgrFlag;
        
      } catch (err) {
        await this.addError(this.properties.siteName,"checkManagerPermission",err.message);
      }
      return;
    }
  
    // **** Function  : checkPowerUserPermission
    // **** Purpose   : is the current user a Power User? check 365 groups.
    // ****
    // ****
    private async checkPowerUserPermission(): Promise<void> {
      this.properties.isDCPowerUser = false;
      this.properties.asmPowerUser = false;
      this.properties.cenPowerUser = false;
      this.properties.cnnPowerUser = false;
      this.properties.empPowerUser = false;
      this.properties.heaPowerUser = false;
      this.properties.dcURL = "";
  
      try {
        const alertHTML: string = '<div class="alert alert-info" role="alert">Checking for Power User Permission - Please wait...</div>';
        const alertContainer: Element = this.domElement.querySelector("#headerBar");
        alertContainer.innerHTML = alertHTML;
  
        graphClient = await this.context.msGraphClientFactory.getClient("3");
        const myGroups = await graphClient.api("/me/transitiveMemberOf/microsoft.graph.group").get();
  
        for (let i = 0; i < myGroups.value.length; i++) {
          // *** check for membership in Team Power User group
          //const groupName = myGroups.value[i].displayName;
          //const position = groupName.search("Power");
          //const groupTeam = groupName.substring(0, position);
          //console.log("myGroup="+groupTeam);
  
          // what happens if a user is in more than one DC Power User group?
          switch (myGroups.value[i].displayName) {
            case "Assessments Document Centre Power Users":
              this.properties.dcURL ="https://" + this.properties.tenantURL[2] + `/sites/asm_dc`;
              this.properties.asmPowerUser = true;
              break;
  
            case "Central Document Centre Power Users":
              this.properties.dcURL ="https://" + this.properties.tenantURL[2] + `/sites/cen_dc`;
              this.properties.cenPowerUser = true;
              break;
  
            case "Connect Document Centre Power Users":
              this.properties.dcURL ="https://" + this.properties.tenantURL[2] + `/sites/cnn_dc`;
              this.properties.cnnPowerUser = true;
              break;
  
            case "Employability Document Centre Power Users":
              this.properties.dcURL ="https://" + this.properties.tenantURL[2] + `/sites/emp_dc`;
              this.properties.empPowerUser = true;
              break;
  
            case "Health Document Centre Power Users":
              this.properties.dcURL ="https://" + this.properties.tenantURL[2] + `/sites/hea_dc`;
              this.properties.heaPowerUser = true;
              break;
          }
        }
  
        if (
          this.properties.asmPowerUser &&
          this.properties.divisionTitle === "Assessments"
        ) {this.properties.isDCPowerUser = true;}
  
        if (
          this.properties.cenPowerUser &&
          this.properties.divisionTitle === "Central"
        ) {this.properties.isDCPowerUser = true;}
  
        if (
          this.properties.cnnPowerUser &&
          this.properties.divisionTitle === "Connect"
        ) {this.properties.isDCPowerUser = true;}
  
        if (
          this.properties.empPowerUser &&
          this.properties.divisionTitle === "Employability"
        ) {this.properties.isDCPowerUser = true;}
  
        if (
          this.properties.heaPowerUser &&
          this.properties.divisionTitle === "Health"
        ) {this.properties.isDCPowerUser = true;}
  
        setTimeout(() => {
          alertContainer.innerHTML = "";
        }, 1000);
        return;
      } catch (err) {
        await this.addError(
          this.properties.siteName,
          "checkPowerUserPermission",
          err.message
        );
      }
    }
  
    // **** Function  : _renderCustomTabsAsync
    // **** Purpose   : set termID based on Divsion
    // ****
    // ****
    private async renderCustomTabsAsync(): Promise<void> {
      const setID = "be84d0a6-e641-4f6d-830e-11e81f13e2f1";
      let termID: string = "";
      let label: string = "";
  
      try {
        switch (this.properties.divisionTitle) {
          case "Assessments":
            termID = "90a0a9eb-bbcc-4693-9674-e56c4d41375f";
            break;
          case "Central":
            termID = "471a563b-a4d9-4ce7-a8e6-4124562b3ace";
            break;
          case "Connect":
            termID = "3532f8fc-4ad2-415c-94ff-c5c7af559996";
            break;
          case "Employability":
            termID = "feb3d3c8-d948-4d3e-b997-a2ea74653b3e";
            break;
          case "Health":
            termID = "c9dfa3b6-c7c6-4e74-a738-0ffe54e1ff5c";
            break;
        }
  
        await this.getCustomTabs(setID, termID).then(async (terms) => {
          for (let x = 0; x < terms.value.length; x++) {
            label = terms.value[x].labels[0].name;
            termID = terms.value[x].id;
  
            if (label === this.properties.siteName) {
              await this.getCustomTabs(setID, termID).then(async (response) => {
                await this._renderCustomTabs(response.value);
                await this.setCustomLibraryListeners();
              });
            }
          }
        });
      } catch (err) {
        await this.addError(this.properties.siteName,"_renderCustomTabsAsync",err.message);
      }
    }
  
    // **** Function  : getCustomTabs
    // **** Purpose   : check termstore for the custom library tabs, for a given setID and termID.
    // ****
    // ****
    private async getCustomTabs(setID: string, termID: string): Promise<any> {
      try {
        const groupID = "4660ef58-779c-4970-bcd7-51773916e8dd";
        const url: string =
          this.context.pageContext.web.absoluteUrl +
          `/_api/v2.1/termStore/groups/${groupID}/sets/${setID}/terms/${termID}/children?select=id,labels`;
        return this.context.spHttpClient
          .get(url, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
            if (response.ok) {
              return response.json();
            }
          });
      } catch (err) {
        await this.addError(this.properties.siteName,"getCustomTabs",err.message);
      }
    }
  
    // **** Function  : _renderCustomTabs
    // **** Purpose   : write the custom library tabs to the DOM
    // ****
    // ****
    private async _renderCustomTabs(items: any): Promise<void> {
      let html: string = "";
      let labelName: string = "";
      let labelID: string = "";
  
      try {
        for (let x = 0; x < items.length; x++) {
          labelName = items[x].labels[0].name;
          labelID = labelName.replace(/\s+/g, "");
  
          this.checkDataAsync("Custom",this.properties.siteName,labelName);
          //console.log("dataFlag="+this.properties.dataFlag);
  
          //if(this.properties.dataFlag){
  
          if (this.properties.isDCPowerUser) {
            html += `<div class="row"> 
                      <button class="btn libraryBtn nav-link mb-1" id="customTab" data-bs-toggle="pill" data-bs-target="#${labelID}" type="button" role="tab" aria-controls="${labelID}" aria-selected="true">
                        <div class="col-1 libraryUploadIcon">
                          <a href="${this.properties.dcURL}/custom/Forms/${this.properties.siteShortName}.aspx" target="_blank">
                            <h5 class="text-white"><i class="bi bi-cloud-arrow-up"></i></h5>
                          </a>
                        </div>
                        <div class="col-10 libraryName"><h6 class="libraryText">${labelName}</h6></div>
                      </button>
                    </div>`;
          } else {
            html += `<div class="row"><button class="btn libraryBtn nav-link text-left mb-1" id="CustomTab" data-bs-toggle="pill" data-bs-target="#Custom" type="button" role="tab" aria-controls="Custom" aria-selected="true">${labelName}</button></div>`;
          }
        }
  
        const listContainer: Element = this.domElement.querySelector("#libraryTabs");
        listContainer.innerHTML += html;
      } catch (err) {
        await this.addError(this.properties.siteName,"_renderCustomTabs",err.message);
      }
      return;
    }
  
    private checkDataAsync(library:string,team:string,category:string):void {
  
      const divisions : string[] = ["Assessments","Central","Connect","Employability","Health"];
  
      for(let x=0;x<divisions.length;x++){
        let division : string = divisions[x];      
        let dcName : string = "";      
  
        switch(division){
          case "Assessments":
            dcName = "asm_dc";
            break;
          case "Central":
            dcName = "cen_dc";
            break;
          case "Connect":
            dcName = "cnn_dc";
            break;
          case "Employability":
            dcName = "emp_dc";
            break;
          case "Health":
            dcName = "hea_dc";
            break;
        }
  
        console.log(division+" "+dcName+" "+library);
      
        this.checkData(dcName,library,team,category)
          .then((response) => {
            console.log(response.value.length);
            if(response.value.length>0){
              this.properties.dataFlag = true; 
            }else{
              this.properties.dataFlag = false;
            }
          })
      }
    }
  
    private checkData(dcName:string,library:string,team:string,category:string):Promise<ISPLists> {
      //this.context.pageContext.web.absoluteUrl +      
      let requestUrl = `https://maximusunitedkingdom.sharepoint.com/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$filter=TaxCatchAll/Term eq '${team}'&$top=10`;
  
      return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        .then((response : SPHttpClientResponse) => {
          return response.json();
        });   
    }
  
      //if(response.ok){
      //  response.json().then((data) => {
          //console.log("requestUrl="+requestUrl);
          //console.log(data.value.length);
      //    console.log("checking dcName="+dcName+" library="+library);
      //    this.properties.numRecords=data.value.length;
  
      //    if(data.value.length>0){              
      //      console.log("data found in "+dcName+" "+library);                                
      //    }else{
      //      console.log("no data found");
      //    }
      //  });
      //}
  
    // **** Function  : getFolders
    // **** Purpose   : fetch the folders for a given library
    // ****
    // ****
    private async getFolders(libraryName: string,tabNum: number,category: string): Promise<void> {
      let fCount = 0;
      let folderContainer: Element;
  
      let folderHTML: string = "";
      let subFolder1HTML : string = "";
      let subFolder2HTML : string = "";
      let subFolder3HTML : string = "";
  
      let folderName: string = "";
      let subFolderName1 : string = "";
      let subFolderName2 : string = "";
      let subFolderName3 : string = "";
  
      let folderPrev: string = "";
      let subFolderPrev1: string = "";
      let subFolderPrev2: string = "";
      let subFolderPrev3: string = "";
  
      let folderNameID : string = ""; // folder id's for event listeners on button click
      let subFolderName1ID: string = "";
      let subFolderName2ID: string = "";
      let subFolderName3ID: string = "";
      
      let folderID : string = ""; // accordion id's for folders
      let subFolder01ID : string = "";
      let subFolder02ID : string = "";
      //let subFolder03ID : string = "";
  
      this.properties.dataResults = [];
      this.properties.folderArray = [];
      this.properties.subFolder1Array = []; // used for the Folder EventListeners
      this.properties.subFolder2Array = [];
      this.properties.subFolder3Array = [];
      this.properties.sf01IDArray=[]; // used for the Folder HTML code
      this.properties.sf02IDArray=[];
      this.properties.sf03IDArray=[];
      this.properties.libraryName = libraryName;
  
      try {
        switch (libraryName) {
          case "Policies":
            folderContainer = this.domElement.querySelector("#policiesFolders");
            folderContainer.innerHTML = "";
            break;
          case "Procedures":
            folderContainer = this.domElement.querySelector("#proceduresFolders");
            folderContainer.innerHTML = "";
            break;
          case "Guides":
            folderContainer = this.domElement.querySelector("#guidesFolders");
            folderContainer.innerHTML = "";
            break;
          case "Forms":
            folderContainer = this.domElement.querySelector("#formsFolders");
            folderContainer.innerHTML = "";
            break;
          case "General":
            folderContainer = this.domElement.querySelector("#generalFolders");
            folderContainer.innerHTML = "";
            break;
          case "Management":
            folderContainer = this.domElement.querySelector("#managementFolders");
            folderContainer.innerHTML = "";
            break;
          case "Custom":
            folderContainer = this.domElement.querySelector("#customFolders");
            folderContainer.innerHTML = "";
            break;
        }
  
        await this.get_Data(libraryName, category).then(async (results) => {
          console.log("results in getFolders");
          console.log(results);
          console.log(results.length);
          console.log(libraryName);
  
          if (results.length > 0) {
            for (let x = 0; x < results.length; x++) {
              const fileStatus: string = results[x].FieldValuesAsText.OData__x005f_ModerationStatus;
  
              if (fileStatus === "Approved" || (fileStatus === "Draft" && this.properties.isDCPowerUser)) {
                
                if(results[x].Knowledge_Folder !== null) {
  
                  folderName = results[x].Knowledge_Folder.Label;
                  folderNameID = folderName.replace(/\s+/g, "");
  
                  // *** Parent Folder ID
                  folderID = "dcTab" + tabNum + "-Folder" + fCount;
                  
                  // *** check is a new folder, if so create new folder string and add to DOM
                  if (folderName !== folderPrev) {
                    console.log("pass x="+x+" f="+folderName+" fprev="+folderPrev);
                    folderHTML += `<div class="accordion mt-1" id="accordionPF${x}">
                                    <div class="accordion-item">
                                      <h2 class="accordion-header" id="folder_${folderNameID}">
                                        <button class="btn btn-primary folderBtn" id="${folderNameID}" type="button" data-bs-toggle="collapse" data-bs-target="#${folderID}" aria-expanded="true" aria-controls="accordionPF${x}">
                                          <i class="bi bi-folder2"></i>
                                          <span class="text-white ms-1">${folderName}</span>                    
                                        </button>
                                      </h2>
                                      <div id="${folderID}" class="accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordionPF${x}">
                                        <div class="accordion-body" id="${folderNameID}Group">
                                        </div>
                                      </div>                                  
                                    </div>
                                  </div>`;
                    fCount++;
                    folderPrev = folderName;                 
                  }                                      
                  this.properties.folderArray.push(folderName);
  
                  if (results[x].Knowledge_SubFolder !== null) {
                    subFolderName1 = results[x].Knowledge_SubFolder.Label;
                    subFolderName1ID = subFolderName1.replace(/\s+/g, "");
                    subFolder01ID = folderID + "-Sub01";
                    this.properties.subFolder1Array.push(subFolderName1);
  
                    if(subFolderName1 !== subFolderPrev1){  
                      console.log("pass x="+x+" sf1="+subFolderName1+" sf1prev="+subFolderPrev1+" sfName2_ID="+subFolderName1ID);
                      this.properties.sf01IDArray.push(folderNameID);                    
                      subFolder1HTML += `<div class="accordion" id="accordion${subFolder01ID}">
                                            <div class="accordion-item">                
                                              <h2 class="accordion-header" id="folder_${subFolderName1ID}">
                                                <button class="btn btn-primary sub1FolderBtn" id="${subFolderName1ID} type="button" data-bs-toggle="collapse" data-bs-target="#${subFolder01ID}" aria-expanded="true" aria-controls="collapseSF3-1">
                                                  <i class="bi bi-folder2"></i>
                                                  <span class="text-white ms-1">${subFolderName1}</span>
                                                </button>
                                              </h2>
                                              <div id="${subFolder01ID}" class="accordion-collapse collapse" aria-labelledby="headingSF3-1" data-bs-parent="accordion${subFolder01ID}">
                                                <div class="accordion-body subFolder" id="${subFolderName1ID}Group">
                                                </div>
                                              </div>
                                            </div>
                                          </div>`;                                   
    
                      subFolderPrev1 = subFolderName1;
                    }                                 
                  
                    if (results[x].Knowledge_SubFolder2 !== null) {
                      subFolderName2 = results[x].Knowledge_SubFolder2.Label;
                      subFolderName2ID = subFolderName2.replace(/\s+/g, "");
                      subFolder02ID = folderID + "-Sub02";
                      this.properties.subFolder2Array.push(subFolderName2);                
                      
                      if(subFolderName2 !== subFolderPrev2){
                        console.log("pass x="+x+" sf2="+subFolderName2+" sf2prev="+subFolderPrev2+" sfName2_ID="+subFolderName2ID);
                        this.properties.sf02IDArray.push(subFolderName1ID);
                        subFolder2HTML += `<div class="accordion" id="accordion${subFolder02ID}">
                                            <div class="accordion-item">                 
                                              <h2 class="accordion-header" id="folder_${subFolderName2ID}">
                                                <button class="btn btn-primary sub1FolderBtn" id="${subFolderName2ID}" type="button" data-bs-toggle="collapse" data-bs-target="#${subFolder02ID}" aria-expanded="true" aria-controls="collapseSF3-2">
                                                  <i class="bi bi-folder2"></i>
                                                  <span class="text-white ms-1">${subFolderName2}</span>
                                                </button>
                                              </h2>
                                              <div id="${subFolder02ID}" class="accordion-collapse collapse" aria-labelledby="headingSF-2" data-bs-parent="accordion${subFolder02ID}">
                                                <div class="accordion-body subFolder" id="${subFolderName2ID}Group">
                                                </div>
                                              </div>
                                            </div>
                                          </div>`;
  
                        subFolderPrev2 = subFolderName2;
                      }
  
                      if (results[x].Knowledge_SubFolder3 !== null) {
                        subFolderName3 = results[x].Knowledge_SubFolder3.Label;
                        subFolderName3ID = subFolderName3.replace(/\s+/g, "");
                        //subFolder03ID = folderID + "-Sub03";
  
                        if(subFolderName3 !== subFolderPrev3){
                          console.log("pass x="+x+" sf3="+subFolderName3+" sf3prev="+subFolderPrev3+" sfName3_ID="+subFolderName3ID);
                          this.properties.sf03IDArray.push(subFolderName2ID);
                          subFolder3HTML += `<h2 class="accordion-header" id="folder_${subFolderName3ID}">
                                              <button class="btn btn-primary sub3FolderBtn" type="button">
                                                <i class="bi bi-folder2"></i>
                                                <span class="text-white ms-1">${subFolderName3}</span>                   
                                              </button>
                                            </h2>`
                          subFolderPrev3 = subFolderName3;                  
                        }
                        this.properties.subFolder3Array.push(subFolderName3);
                      }
                    } // end of subfolder 2 check
                  }  // end of subfolder 1 check                                
                } // end of folder check
              } // end of file status check
            } // end of for loop
  
            folderContainer.innerHTML += folderHTML;
  
            console.log("F_id=" + this.properties.sf01IDArray);
            console.log("SF1_id=" + this.properties.sf02IDArray);
            console.log("SF2_id=" + this.properties.sf03IDArray);
    
            if(this.properties.sf01IDArray !== undefined){
              for(let x=0;x<this.properties.sf01IDArray.length;x++){
                if(this.properties.sf01IDArray[x]!==this.properties.sf01IDArray[x-1]){
                  $('#'+this.properties.sf01IDArray[x]+'Group').append(subFolder1HTML);
                  const elem = document.querySelector("#"+this.properties.sf01IDArray[x]);
                  elem.classList.add('accordion-button');
                }
              }
            }
            
            if(this.properties.sf02IDArray !== undefined){
              for(let x=0;x<this.properties.sf02IDArray.length;x++){
                if(this.properties.sf02IDArray[x]!==this.properties.sf02IDArray[x-1]){
                  $('#'+this.properties.sf02IDArray[x]+'Group').append(subFolder2HTML);
                  const elem = document.querySelector("#"+this.properties.sf02IDArray[x]);
                  elem.classList.add('accordion-button');
                }
              }
            }
  
            if(this.properties.sf03IDArray !== undefined){
              for(let x=0;x<this.properties.sf03IDArray.length;x++){
                if(this.properties.sf03IDArray[x]!==this.properties.sf03IDArray[x-1]){
                  $('#'+this.properties.sf03IDArray[x]+'Group').append(subFolder3HTML);
                  //const elem = document.querySelector("#"+this.properties.sf02IDArray[x]);
                  //elem.classList.add('accordion-button');
                }
              }           
            }
  
          } else {
            alertHTML = `<div class="alert alert-warning" role="alert">There are no documents for ${this.properties.siteName} 
                          ${libraryName === "Custom" ? ` & ${category} category` : `in the ${libraryName}`} 
                          ${this.properties.divisionName === "All" ? "libraries" : "library"}
                        </div>`;
            alertContainer = this.domElement.querySelector("#headerBar");
            alertContainer.innerHTML = alertHTML;
          }
          console.log("F=" + this.properties.folderArray);
          console.log("SF1=" + this.properties.subFolder1Array);
          console.log("SF2=" + this.properties.subFolder2Array);
          console.log("SF3=" + this.properties.subFolder3Array);
          await this.setFolderListeners();
        });
  
        // *** remove alert after 1.5 seconds
        const alertText = alertContainer.innerHTML;
        if (alertText !== "") {
          setTimeout(() => {
            alertContainer.innerHTML = "";
          }, 1500);
        }
        //await this.setFolderListeners();
      } catch (err) {
        console.log("getFolders : ",err.lineNumber);
        await this.addError(this.properties.siteName, "getFolders", err.message);
      }
      return;
    }
  
    // **** Function  : getData
    // **** Purpose   : fetch the data from each Divisional Document Centre, add to the data results array.
    // ****
    // ****
    private async get_Data(libraryName: string,category: string): Promise<any[]> {
  
      const asmlist = this.asmDC.lists.getByTitle(libraryName);
      const cenlist = this.cenDC.lists.getByTitle(libraryName);
      const cnnlist = this.cnnDC.lists.getByTitle(libraryName);
      const emplist = this.empDC.lists.getByTitle(libraryName);
      const healist = this.heaDC.lists.getByTitle(libraryName);
  
      let view: string = "";
      this.properties.dataResults = [];
      //this.properties.folderArray=[];
  
      alertHTML = `<div class="alert alert-info" role="alert">
                    Fetching data for ${this.properties.siteName} 
                    team from the ${libraryName} ${this.properties.divisionName === "All" ? "libraries" : "library"} 
                    ${libraryName === "Custom" ? ` for ${category} category` : ``}
                  </div>`;
      alertContainer = this.domElement.querySelector("#headerBar");
      alertContainer.innerHTML = alertHTML;
  
      try {
        if (category === "") {
          view =
            "<View><Query>" +
              "<Where>" +
                "<Or>" +
                  "<Eq>" +
                    '<FieldRef Name="Knowledge_Team"/>' +
                    '<Value Type="TaxonomyFieldType">' +this.properties.siteName + "</Value>" +
                  "</Eq>" +
                  "<Eq>" +
                    '<FieldRef Name="Knowledge_SharedWith"/>' +
                    '<Value Type="TaxonomyFieldTypeMulti">' + this.properties.siteName +"</Value>" +
                  "</Eq>" +
                "</Or>" +
              "</Where>" +
              "<OrderBy>" +
                '<FieldRef Name="Knowledge_Folder" Ascending="TRUE" />' +
                '<FieldRef Name="Knowledge_SubFolder" Ascending="TRUE" />'+
                '<FieldRef Name="Knowledge_SubFolder2" Ascending="TRUE" />'+
                '<FieldRef Name="Knowledge_SubFolder3" Ascending="TRUE" />'+
                '<FieldRef Name="LinkFilename" Ascending="TRUE" />' +
              "</OrderBy>" +
            "</Query></View>";
        } else {
          view =
            "<View><Query>" +
              "<Where>" +
                "<And>" +
                  "<Eq>" +
                    '<FieldRef Name="Knowledge_Category"/>' +
                    '<Value Type="TaxonomyFieldType">' + category + '</Value>' +
                  "</Eq>" +
                  "<Or>" +
                    "<Eq>" +
                      '<FieldRef Name="Knowledge_Team"/>' +
                      '<Value Type="TaxonomyFieldType">' + this.properties.siteName + '</Value>' +
                    "</Eq>" +
                    "<Eq>" +
                      '<FieldRef Name="Knowledge_SharedWith"/>' +
                      '<Value Type="TaxonomyFieldTypeMulti">' + this.properties.siteName + '</Value>' +
                    "</Eq>" +
                  "</Or>" +
                "</And>" +
              "</Where>" +
              "<OrderBy>" +
                '<FieldRef Name="Knowledge_Folder" Ascending="TRUE" />' +
                '<FieldRef Name="Knowledge_SubFolder" Ascending="TRUE" />'+
                '<FieldRef Name="Knowledge_SubFolder2" Ascending="TRUE" />'+
                '<FieldRef Name="Knowledge_SubFolder3" Ascending="TRUE" />'+
                '<FieldRef Name="LinkFilename" Ascending="TRUE" />' +
              "</OrderBy>" +
            "</Query></View>";
        }
  
        // *** check against current site / DC (divisionTitle) and also if a Property Pane item has been selected (divisionName)
        //console.log("division title=" +this.properties.divisionTitle +" division name " +this.properties.divisionName);
        //console.log("view="+view);
  
        // *** get data from Assessments Document Centre and add to dataResults array
        if (
          this.properties.divisionTitle === "Assessments" ||
          this.properties.divisionName === "ASM" ||
          this.properties.divisionName === "All"
        ) {
          await asmlist
            .getItemsByCAMLQuery(
              { ViewXml: view },
              "FieldValuesAsText/FileRef",
              "FieldValueAsText/FileLeafRef"
            )
            .then((asm_Results: string | any[]) => {
              if (asm_Results.length > 0) {
                for (let c = 0; c < asm_Results.length; c++) {
                  this.properties.dataResults.push(asm_Results[c]);
                }
              }
              console.log("ASM DC Results");
              console.log(asm_Results);
            });
        }
  
        // *** get data from Central Document Centre and add to dataResults array
        if (
          this.properties.divisionTitle === "Central" ||
          this.properties.divisionName === "CEN" ||
          this.properties.divisionName === "All"
        ) {
          await cenlist
            .getItemsByCAMLQuery(
              { ViewXml: view },
              "FieldValuesAsText/FileRef",
              "FieldValueAsText/FileLeafRef"
            )
            .then((cen_Results: string | any[]) => {
              if (cen_Results.length > 0) {
                for (let c = 0; c < cen_Results.length; c++) {
                  this.properties.dataResults.push(cen_Results[c]);
                }
              }
              console.log("CEN DC Results");
              console.log(cen_Results);
            });
        }
  
        // *** get data from Connect Document Centre and add to dataResults array
        if (
          this.properties.divisionTitle === "Connect" ||
          this.properties.divisionName === "CNN" ||
          this.properties.divisionName === "All"
        ) {
          await cnnlist
            .getItemsByCAMLQuery(
              { ViewXml: view },
              "FieldValuesAsText/FileRef",
              "FieldValueAsText/FileLeafRef"
            )
            .then((cnn_Results: string | any[]) => {
              if (cnn_Results.length > 0) {
                for (let c = 0; c < cnn_Results.length; c++) {
                  this.properties.dataResults.push(cnn_Results[c]);
                }
              }
              console.log("CNN DC Results");
              console.log(cnn_Results);
            });
        }
  
        // *** get data from Employability Document Centre and add to dataResults array
        if (
          this.properties.divisionTitle === "Empoyability" ||
          this.properties.divisionName === "EMP" ||
          this.properties.divisionName === "All"
        ) {
          await emplist
            .getItemsByCAMLQuery(
              { ViewXml: view },
              "FieldValuesAsText/FileRef",
              "FieldValueAsText/FileLeafRef"
            )
            .then((emp_Results: string | any[]) => {
              if (emp_Results.length > 0) {
                for (let c = 0; c < emp_Results.length; c++) {
                  this.properties.dataResults.push(emp_Results[c]);
                }
              }
              console.log("EMP DC Results");
              console.log(emp_Results);
            });
        }
  
        // *** get data from Health Document Centre and add to dataResults array
        if (
          this.properties.divisionTitle === "Health" ||
          this.properties.divisionName === "HEA" ||
          this.properties.divisionName === "All"
        ) {
          await healist
            .getItemsByCAMLQuery(
              { ViewXml: view },
              "FieldValuesAsText/FileRef",
              "FieldValueAsText/FileLeafRef"
            )
            .then((hea_Results: string | any[]) => {
              if (hea_Results.length > 0) {
                for (let c = 0; c < hea_Results.length; c++) {
                  this.properties.dataResults.push(hea_Results[c]);
                }
              }
              console.log("HEA DC Results");
              console.log(hea_Results);
            });
        }
      } catch (err) {
        await this.addError(this.properties.siteName, "getData", err.message);
      }
      return this.properties.dataResults;
    }
  
    // **** Function  : getFiles
    // **** Purpose   : To filter the file data results array where the Folder or Subfolder exist in the file object,
    // ****             then add the file object to the Folder or SubFolder Array as required.
    // ****             Passing this array to the ShowFiles function for processing to DOM.
  
    private getFiles(folderID: string, folderName: string): Promise<void> {
  
      const libraryName: string = this.properties.libraryName.toLowerCase();
  
      // *** clear the filesArray
      this.properties.filesArray = [];
      
      // *** clear the files container of any HTML
      const filesContainer: Element = this.domElement.querySelector("#" + libraryName + "Files");
      if (filesContainer.innerHTML !== null || filesContainer.innerHTML !== "") {filesContainer.innerHTML = "";}
  
      console.log("getFiles called");
      console.log("library=" +libraryName +" folderID=" +folderID +" folderName=" +folderName);
  
      try{
  
        if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
          document
            .querySelector("#" + libraryName + "Folders button.active")
            .classList.remove("active");
        }
  
        const elem = document.querySelector("#folder_" + folderID);
        const elemCSS = elem.classList.contains('active');
  
        if( elemCSS !== true){
          alertHTML = `<div class="alert alert-info" role="alert">Fetching files for ${folderName} folder.</div>`;
          alertContainer = this.domElement.querySelector('#headerBar');
          alertContainer.innerHTML = alertHTML;
        }
        elem.classList.add('active');
  
        const folderFiles = this.properties.dataResults.filter(function (item: any) {
          if (item.Knowledge_Folder.Label === folderName && item.Knowledge_SubFolder === null) {
            return item;
          }
        });
  
        const subFolder1Files = this.properties.dataResults.filter(function (item: any) {
          if (item.Knowledge_SubFolder !== null && item.Knowledge_SubFolder2 === null) {
            if (item.Knowledge_SubFolder.Label === folderName) {
              return item;
            }
          }
        });
  
        const subFolder2Files = this.properties.dataResults.filter(function(item){
          if(item.Knowledge_SubFolder2 !== null && item.Knowledge_SubFolder3 === null){
            if(item.Knowledge_SubFolder2.Label===folderName){
              return item;
            }
          }
        });
  
        const subFolder3Files = this.properties.dataResults.filter(function(item){
          if(item.Knowledge_SubFolder3 !== null){
            if(item.Knowledge_SubFolder3.Label===folderName){
              return item;
            }
          }
        });
  
        if (folderFiles.length > 0 && subFolder1Files.length === 0) {
          this.properties.filesArray=folderFiles;
          this.showFiles(libraryName);
        }
  
        if (subFolder1Files.length > 0 && subFolder2Files.length === 0) {
          this.properties.filesArray=subFolder1Files;
          this.showFiles(libraryName);
        }
      
        if (subFolder2Files.length > 0 && subFolder3Files.length === 0) {
          this.properties.filesArray=subFolder2Files;
          this.showFiles(libraryName);
        }
  
        if (subFolder3Files.length > 0) {
          this.properties.filesArray=subFolder3Files;
          this.showFiles(libraryName);
        }
       
        if (alertContainer.innerHTML !== null) {
          setTimeout(() => {
            alertContainer.innerHTML = "";
          }, 2000);
        }
      }
        catch(err){
        this.addError(this.properties.siteName, "getFiles", err.message);
      }
      this.setFileListeners();
      return;
    }
  
    // **** Function  : showFiles
    // **** Purpose   : write the file array contents to the DOM
    // ****
    // ****
    private async showFiles(libraryName: string): Promise<void> {
      const filesContainer: Element = this.domElement.querySelector("#" + libraryName + "Files");
      let fileHTML: string = "";
      let powerUserTools: string = "";
      let draftText: string = "";
      let fileTypeIcon: string = "";
  
      try {
        if (filesContainer.innerHTML !== null) {
          filesContainer.innerHTML = "";
        }
  
        console.log("File Array");
        console.log(this.properties.filesArray);
  
        for (let x = 0; x < this.properties.filesArray.length; x++) {
          const fileStatus: string = this.properties.filesArray[x].FieldValuesAsText.OData__x005f_ModerationStatus;
          const fileType: string = this.properties.filesArray[x].FieldValuesAsText.File_x005f_x0020_x005f_Type;
          const fileName: string = this.properties.filesArray[x].FieldValuesAsText.FileLeafRef;
          const fileURL: string = this.properties.filesArray[x].ServerRedirectedEmbedUrl;
          const fileID: string = this.properties.filesArray[x].ID;
          //let fileGUID : string=items[x].GUID;
          //console.log(fileGUID);
          //console.log(fileID);
  
          switch (fileType) {
            case "pdf":
              fileTypeIcon = "bi bi-file-earmark-pdf text-danger";
              break;
            case "doc":
            case "docx":
            case "dotx":
              fileTypeIcon = "bi bi-file-earmark-word text-primary";
              break;
            case "xls":
            case "xlsx":
            case "xlsm":
            case "xltx":
              fileTypeIcon = "bi bi-file-earmark-excel text-sucess";
              break;
            case "ppt":
            case "pptx":
            case "potx":
              fileTypeIcon = "bi bi-file-earmark-ppt text-warning";
              break;
            case "xsn":
              fileTypeIcon = "bi bi-file-earmark-excel text-secondary";
              break;
            case "msg":
              fileTypeIcon = "bi bi-envelope-at text-info";
              break;
            case "zip":
            case "rar":
              fileTypeIcon = "bi bi-file-earmark-zip text-danger";
              break;
            default:
              fileTypeIcon = "bi bi-file-earmark text-dark";
              break;
          }
  
          if (this.properties.isDCPowerUser) {
            switch (fileStatus) {
              case "Approved":
                powerUserTools = `<a class="docDelete" id="doc${fileID}Delete" href="#" title="Delete Document"><i class="bi bi-trash"></i></a >`;
                draftText = "";
                break;
              case "Draft":
                powerUserTools = `<a class="docDelete" id="doc${fileID}Delete" href="#" title="Delete Document"><i class="bi bi-trash"></i></a>
                                <a class="docPublish" id="doc${fileID}Publish" href="#" title="Publish Draft Document"><i class="bi bi-file-check"></i></a>`;
                draftText = '<i class="text-danger">draft</i>';
                break;
              default:
                powerUserTools = "";
                draftText = "";
                break;
            }
          }
  
          fileHTML += `<div class="row fileRow" id="fileRow${fileID}">
                        <div class="col-2">
                          ${powerUserTools}
                          <i class="${fileTypeIcon}" title="file icon"></i>
                          <a class="docCopyLink" id="doc${fileID}Link" href="#" title="Copy File Link to Clipboard"><i class="bi bi-link"></i></a>
                        </div>
                        <div class="col-10">
                          ${draftText}
                          <a href="${fileURL}" title="file name" target="_blank"><p>${fileName}</p></a>
                        </div>
                      </div>`;
          filesContainer.innerHTML = fileHTML;
        }
      } catch (err) {
        await this.addError(this.properties.siteName, "showFiles", err.message);
      }
      return;
    }
  
    // **** Function  : deleteFile
    // **** Purpose   : delete a file when the button is clicked
    // ****
    // ****
    private async deleteFile(fileRelativeURL: string,fileName: string,fileID: string): Promise<void> {
      let alertHTML: string = "";
      alertContainer = this.domElement.querySelector("#headerBar");
  
      //this.properties.dcURL ="https://maximusunitedkingdom.sharepoint.com/sites/cen_dc";
      //https://maximusunitedkingdom.sharepoint.com/sites/cen_dc/_api/web/GetFileByServerRelativeUrl('/sites/cen_dc/Policies/test%20doc%2003.xlsx')/etag
  
      const url: string = this.properties.dcURL + `/_api/web/GetFileByServerRelativeUrl('${fileRelativeURL}')/recycle`;
  
      const deleteHeader = {
        Authorization: "Bearer ",
        "If-MATCH": "*",
        "Content-type": "application/json;odata=verbose",
        accept: "application/json;odata=verbose",
        "odata-version": "3.0",
        "X-HTTP-Method": "DELETE",
      };
  
      const options: ISPHttpClientOptions = { headers: deleteHeader };
  
      await this.context.spHttpClient
        .post(url, SPHttpClient.configurations.v1, options)
        .then(async (response: SPHttpClientResponse) => {
          try {
            if (response.ok) {
              alertHTML = `<div class="alert alert-info" role="alert">Deleting Document ${fileName} is being recycled, Please Wait...</div>`;
              document.getElementById("fileRow" + fileID).remove();
              for (let c = 0; c < this.properties.dataResults.length; c++) {
                const fileArrayItem = this.properties.dataResults[c].ID;
                if (fileArrayItem === fileID) {
                  const index = this.properties.dataResults.indexOf(
                    this.properties.dataResults[c]
                  );
                  if (index > -1) {
                    this.properties.dataResults.splice(index, 1);
                  }
                }
              }
            } else {
              alertHTML = `<div class="alert alert-danger" role="alert">Delete Request Failed as ${fileName} could not be deleted, please ensure it is checked in.</div>`;
            }
          } catch (err) {
            await this.addError(this.properties.siteName, "deleteFile", err.message);
          }
        });
  
      alertContainer.innerHTML = alertHTML;
      setTimeout(() => {
        alertContainer.innerHTML = "";
      }, 1500);
  
      return;
    }
  
    // **** Function  : copyLink
    // **** Purpose   : copy a file link to the clipboard when a button is clicked
    // ****
    // ****
    private async copyLink(fileURL: string, fileName: string): Promise<void> {
      let alertHTML: string = "";
      alertContainer = this.domElement.querySelector("#headerBar");
  
      try {
        await navigator.clipboard.writeText(fileURL);
        alertHTML = `<div class="alert alert-info" role="alert">URL link for ${fileName} has been copied to the clipboard</div>`;
      } catch (err) {
        alertHTML = `<div class="alert alert-danger" role="alert">Could not copy URL link for ${fileName} to the clipboard</div>`;
        await this.addError(this.properties.siteName, "copyLink", err.message);
      }
  
      alertContainer.innerHTML = alertHTML;
      setTimeout(() => {
        alertContainer.innerHTML = "";
      }, 1500);
      return;
    }
  
    // **** Function  : publishFile
    // **** Purpose   : publish a file when the button is clicked
    // ****
    // ****
    private async publishFile(fileRelativeURL: string,fileName: string,fileID: string): Promise<void> {
      let alertHTML: string = "";
      const filesContainer: Element = this.domElement.querySelector("#fileRow" + fileID);
      let fileHTML: string = "";
      let fileTypeIcon: string = "";
      let powerUserTools: string = "";
      let draftText: string = "";    
      const url: string =this.properties.dcURL +`/_api/web/GetFileByServerRelativeUrl('${fileRelativeURL}')/Publish()`;
      const publishHeader = {
          Authorization: "Bearer ",
          "If-MATCH": "*",
          "Content-type": "application/json;odata=verbose",
          accept: "application/json;odata=verbose",
          "odata-version": "3.0",
          "X-HTTP-Method": "DELETE",
      };
      const options: ISPHttpClientOptions = { headers: publishHeader };
  
      alertContainer = this.domElement.querySelector("#headerBar");
      this.properties.dcURL ="https://maximusunitedkingdom.sharepoint.com/sites/cen_dc";
  
      await this.context.spHttpClient
        .post(url, SPHttpClient.configurations.v1, options)
        .then(async (response: SPHttpClientResponse) => {
          try {
            if (response.ok) {
              alertHTML = `<div class="alert alert-info" role="alert">Publishing Document ${fileName} to major version, Please Wait...</div>`;
              filesContainer.innerHTML = "";
  
              for (let c = 0; c < this.properties.dataResults.length; c++) {
                const fileArrayItem = this.properties.dataResults[c].ID;
  
                if (fileArrayItem === fileID) {
                  const fileType: string = this.properties.dataResults[c].FieldValuesAsText.File_x005f_x0020_x005f_Type;
                  const fileURL: string = this.properties.dataResults[c].ServerRedirectedEmbedUrl;
  
                  switch (fileType) {
                    case "pdf":
                      fileTypeIcon = "bi bi-file-earmark-pdf text-danger";
                      break;
                    case "doc":
                    case "docx":
                    case "dotx":
                      fileTypeIcon = "bi bi-file-earmark-word text-primary";
                      break;
                    case "xls":
                    case "xlsx":
                    case "xlsm":
                    case "xltx":
                      fileTypeIcon = "bi bi-file-earmark-excel text-sucess";
                      break;
                    case "ppt":
                    case "pptx":
                    case "potx":
                      fileTypeIcon = "bi bi-file-earmark-ppt text-warning";
                      break;
                    case "xsn":
                      fileTypeIcon = "bi bi-file-earmark-excel text-secondary";
                      break;
                    case "msg":
                      fileTypeIcon = "bi bi-envelope-at text-info";
                      break;
                    case "zip":
                    case "rar":
                      fileTypeIcon = "bi bi-file-earmark-zip text-danger";
                      break;
                    default:
                      fileTypeIcon = "bi bi-file-earmark text-dark";
                      break;
                  }
                  if (this.properties.isDCPowerUser) {
                    powerUserTools = `<a class="docDelete" id="doc${fileID}Delete" href="#" title="delete document"><i class="bi bi-trash"></i></a >`;
                    draftText = "";
                  }
  
                  fileHTML += `<div class="row fileRow" id="fileRow${fileID}">
                                <div class="col-2">
                                  ${powerUserTools}
                                  <i class="${fileTypeIcon}"></i>
                                  <a class="docCopyLink" id="doc${fileID}Link" href="#" title="copy link to clipboard"><i class="bi bi-link"></i></a>
                                </div>
                                <div class="col-10">
                                  ${draftText}
                                  <a href="${fileURL}" target="_blank"><p>${fileName}</p></a>
                                </div>
                              </div>`;
                  filesContainer.innerHTML = fileHTML;
                }
              }
            } else {
              alertHTML = `<div class="alert alert-danger" role="alert">Publish Document Request Failed for ${fileName}.</div>`;
            }
          } catch (err) {
            await this.addError(this.properties.siteName, "publishFile", err.message);
          }
        });
      alertContainer.innerHTML = alertHTML;
      setTimeout(() => {
        alertContainer.innerHTML = "";
      }, 1500);
  
      return;
    }
  
    // **** Function  : addError
    // **** Purpose   : write error msg to Document Centre Error log in Developers site.
    // ****
    // ****
    private async addError(siteName: string,funcName: string,errMsg: any): Promise<void> {
      const devSiteURL ="https://maximusunitedkingdom.sharepoint.com/teams/developers";
      const addHeader = {
          Authorization: "Bearer ",
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
      };
      const options = {
        headers: addHeader,
        body: JSON.stringify({
          Title: funcName,
          SiteName: siteName,
          Error: errMsg,
        }),
      };
      await this.context.spHttpClient
        .post(
          devSiteURL + "/_api/web/lists/getbytitle('Document Centre Errors')/Items",
          SPHttpClient.configurations.v1,
          options
        )
        //.then(response => {});
      return;
    }
  
    // **** Function  : setLibraryListeners
    // **** Purpose   : set the event listeners for each library element, to fetch the folders when the button is clicked
    // ****
    // ****
    private async setLibraryListeners(): Promise<void> {
      try {
        // *** event listeners for main document libraries
        document
          .getElementById("PoliciesTab")
          .addEventListener("click", (_e: Event) =>
            this.getFolders("Policies", 1, "")
          );
        document
          .getElementById("ProceduresTab")
          .addEventListener("click", (_e: Event) =>
            this.getFolders("Procedures", 2, "")
          );
        document
          .getElementById("GuidesTab")
          .addEventListener("click", (_e: Event) =>
            this.getFolders("Guides", 3, "")
          );
        document
          .getElementById("FormsTab")
          .addEventListener("click", (_e: Event) =>
            this.getFolders("Forms", 4, "")
          );
        document
          .getElementById("GeneralTab")
          .addEventListener("click", (_e: Event) =>
            this.getFolders("General", 5, "")
          );
  
        // *** event listener for management library
        if (this.properties.isManager) {
          document
            .getElementById("ManagementTab")
            .addEventListener("click", (_e: Event) =>
              this.getFolders("Management", 6, "")
            );
        }
      } catch (err) {
        await this.addError(this.properties.siteName, "setLibraryListeners", err.message);
      }
      return;
    }
  
    // **** Function  : setCustomLibraryListeners
    // **** Purpose   : set the event listeners for each custom library element, to fetch the folders when the button is clicked
    // ****
    // ****
    private async setCustomLibraryListeners(): Promise<void> {
      // *** event listeners for custom library category tabs
      try {
        const tabBtns = this.domElement.querySelectorAll("#CustomTab");
        for (let x = 0; x < tabBtns.length; x++) {
          const customTab = tabBtns[x].innerHTML;
          tabBtns[x].addEventListener("click", (_e: Event) =>
            this.getFolders("Custom", 7, customTab)
          );
        }
      } catch (err) {
        await this.addError(
          this.properties.siteName,
          "setCustomLibraryListeners",
          err
        );
      }
      return;
    }
  
    // **** Function  : setFolderListeners
    // **** Purpose   : set the event listeners for each folder or subfolder element, to fetch the files when the button is clicked
    // ****
    // ****
    private async setFolderListeners(): Promise<void> {
      console.log("setFolderListeners called ");
      //this.properties.getFilesCallFlag = false;
      //console.log("setfolderlisteners callflag="+this.properties.getFilesCallFlag);
      try {
        // *** event listeners for parent folders
              
        if (this.properties.folderArray.length > 0) {
          for (let x = 0; x < this.properties.folderArray.length; x++) {
            //if()
            const folderNameID = this.properties.folderArray[x].replace(/\s+/g, "");
            //console.log("libraryName="+this.properties.libraryName+" folderIDTemp="+this.properties.folderArray[x].replace(/\s+/g,"")+" folderNameTemp="+this.properties.folderArray[x]);
            document
              .getElementById("folder_" + folderNameID)
              .addEventListener("click", (_e: Event) =>
                this.getFiles(folderNameID, this.properties.folderArray[x])
              );
          }
        }
       
        if (this.properties.subFolder1Array.length > 0) {
          for (let x = 0; x < this.properties.subFolder1Array.length; x++) {
            const subFolder1NameID = this.properties.subFolder1Array[x].replace(/\s+/g,"");
            //console.log("libraryName="+this.properties.libraryName+" folderIDTemp="+this.properties.subFolder1Array[x].replace(/\s+/g,"")+" folderNameTemp="+this.properties.subFolder1Array[x]);
            document
              .getElementById("folder_" + subFolder1NameID)
              .addEventListener("click", (_e: Event) =>
                this.getFiles(
                  subFolder1NameID,
                  this.properties.subFolder1Array[x]
                )
              );
          }
        }
  
        if (this.properties.subFolder2Array.length > 0) {
          for (let x = 0; x < this.properties.subFolder2Array.length; x++) {
            const subFolder2NameID = this.properties.subFolder2Array[x].replace(/\s+/g,"");
            //console.log("libraryName="+this.properties.libraryName+" folderIDTemp="+this.properties.subFolder2Array[x].replace(/\s+/g,"")+" folderNameTemp="+this.properties.subFolder2Array[x]);
            document
              .getElementById("folder_" + subFolder2NameID)
              .addEventListener("click", (_e: Event) =>
                this.getFiles(
                  subFolder2NameID,
                  this.properties.subFolder2Array[x]
                )
              );
          }
        }
  
        if (this.properties.subFolder3Array.length > 0) {
          for (let x = 0; x < this.properties.subFolder3Array.length; x++) {
            const subFolder3NameID = this.properties.subFolder3Array[x].replace(/\s+/g,"");
            console.log("libraryName=" +this.properties.libraryName +" folderIDTemp=" +this.properties.subFolder3Array[x].replace(/\s+/g,"") +" folderNameTemp=" +this.properties.subFolder3Array[x]);
            document
              .getElementById("folder_" + subFolder3NameID)
              .addEventListener("click", (_e: Event) =>
                this.getFiles(
                  subFolder3NameID,
                  this.properties.subFolder3Array[x]
                )
              );
          }
  
      } catch (err) {
        //console.log("setFolderListeners", err.message);
        await this.addError(this.properties.siteName, "setFolderListeners", err.message);
      }
      return;
    }
  
    // **** Function  : setFileListeners
    // **** Purpose   : set the event listeners for each file element, call the function when the icon is clicked
    // ****
    // ****
    private async setFileListeners(): Promise<void> {
      try {
        // *** event listeners for copy link icon
        for (let x = 0; x < this.properties.filesArray.length; x++) {
          const fileURL: string =this.properties.filesArray[x].ServerRedirectedEmbedUrl;
          const fileRefURL: string =this.properties.filesArray[x].FieldValuesAsText.FileRef;
          const fileID: string = this.properties.filesArray[x].ID;
          const fileName: string =this.properties.filesArray[x].FieldValuesAsText.FileLeafRef;
          const fileStatus: string =this.properties.filesArray[x].FieldValuesAsText.OData__x005f_ModerationStatus;
          document
            .getElementById("doc" + fileID + "Link")
            .addEventListener("click", (_e: Event) =>
              this.copyLink(fileURL, fileName)
            );
  
          // *** event listeners for file delete icon
          if (this.properties.isDCPowerUser) {
            document
              .getElementById("doc" + fileID + "Delete")
              .addEventListener("click", (_e: Event) =>
                this.deleteFile(fileRefURL, fileName, fileID)
              );
  
            // *** event listeners for file publish icon
            if (fileStatus === "Draft") {
              document
                .getElementById("doc" + fileID + "Publish")
                .addEventListener("click", (_e: Event) =>
                  this.publishFile(fileRefURL, fileName, fileID)
                );
            }
          }
        }
      } catch (err) {
        await this.addError(this.properties.siteName, "setFileListeners", err.message);
      }
      return;
    }
  
    public async render(): Promise<void> {
      const bootstrapCssURL ="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css";
      const bootstrapIconsCssURL ="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css";
      SPComponentLoader.loadCss(bootstrapCssURL);
      SPComponentLoader.loadCss(bootstrapIconsCssURL);
  
      this.properties.folderArray = [];
  
      // *** get Tenant URL, Division Title, Site Name
      this.properties.URL = this.context.pageContext.web.absoluteUrl;
      this.properties.tenantURL = this.properties.URL.split("/", 5);
      const siteSNArray: any[] = this.properties.URL.split("_", 2);
      this.properties.siteShortName = siteSNArray[1];
      this.properties.siteTitle = this.context.pageContext.web.title;
      this.properties.siteArray = this.properties.siteTitle.split(" - ");
      this.properties.divisionTitle = this.properties.siteArray[0];
      this.properties.siteName = this.properties.siteArray[1];
  
      //${escape(this.context.pageContext.user.displayName)}
  
      this.domElement.innerHTML = `
      <section class="${styles.documentCentre} ${
        !!this.context.sdks.microsoftTeams ? styles.teams : ""
      }">    
        <div class="row titleRow text-white rounded"> 
          ${
            this.properties.isDCPowerUser
              ? `<div class="col-1"><a href="${this.properties.dcURL}" title="open Document Centre" target="_blank"><h3 class="text-white"><i class="bi bi-collection"></i></h3></a></div>`
              : `<div class="col-1"><p>.</div>`
          }
          <div class="col-11">
            <div class="row">
              <div class="col-6 welcomeText">
                <h5 class="text-white">Welcome to ${escape(this.properties.siteTitle)} Documents
                  ${[/*<mgt-person person-query="me"></mgt-person>]}

                </h5>
              </div>
              <div class="col-5 poweruserText">
                <h6>${
                  this.properties.isDCPowerUser
                    ? `<i class="bi bi-person-circle"></i> (Power User)`
                    : ""
                }</h6>
              </div>
            </div>
          </div>
        </div>
        
        <div class="container v-scrollbar">
          <div class="row" style="height:fit-content">
            <div class="col-3"><img src="${require("./assets/Robot_Spin.gif")}" height="50" width="50"/></div>
            <div class="col" id="headerBar"></div>
            ${[/*<div class="col">${this.properties.siteTitle}</div>]}

            </div>
          <div class="row">
            <div class="col-auto libraryContainer">
              <div class="d-flex mt-1 align-items-start">
                <div class="nav flex-column nav-pills me-3 libraryList" id="libraryTabs" role="tablist" aria-orientation="vertical"></div> 
              </div>
            </div>
  
            <div class="col-9 tab-content foldersFilesContainer">            
                <div class="tab-pane fade libraryTab" id="Policies" role="tabpanel" aria-labelledby="policies"> 
                  <div class="row">
                    <div class="col-auto" id="policiesFolders"></div>
                    <div class="col" id="policiesFiles"></div>
                  </div>               
                </div>
                <div class="tab-pane fade libraryTab" id="Procedures" role="tabpanel" aria-labelledby="procedures">
                  <div class="row">
                    <div class="col-auto" id="proceduresFolders"></div>
                    <div class="col" id="proceduresFiles"></div>
                  </div> 
                </div>
                <div class="tab-pane fade libraryTab" id="Guides" role="tabpanel" aria-labelledby="guides">
                  <div class="row">
                    <div class="col-auto" id="guidesFolders"></div>
                    <div class="col" id="guidesFiles"></div>
                  </div> 
                </div>
                <div class="tab-pane fade libraryTab" id="Forms" role="tabpanel" aria-labelledby="forms">
                  <div class="row">
                    <div class="col-auto" id="formsFolders"></div>
                    <div class="col" id="formsFiles"></div>
                  </div> 
                </div>
                <div class="tab-pane fade libraryTab" id="General" role="tabpanel" aria-labelledby="general">
                  <div class="row">
                    <div class="col-auto" id="generalFolders"></div>
                    <div class="col" id="generalFiles"></div>
                  </div> 
                </div>
                <div class="tab-pane fade libraryTab" id="Management" role="tabpanel" aria-labelledby="management">
                  <div class="row">
                    <div class="col-auto" id="managementFolders"></div>
                    <div class="col" id="managementFiles"></div>
                  </div> 
                </div>
                <div class="tab-pane fade libraryTab" id="Custom" role="tabpanel" aria-labelledby="custom">
                  <div class="row">
                    <div class="col-auto" id="customFolders"></div>
                    <div class="col" id="customFiles"></div>
                  </div> 
                </div>              
              
            </div>
          </div>
        </div>
      </section>`;
      
      await this.checkPowerUserPermission();
      await this.getLibraryTabs().then(async () => {
        await this.setLibraryListeners(); 
      });
  
      Log.info('DocumentCentre', 'message', this.context.serviceScope);
      Log.warn('DocumentCentre', 'WARNING message', this.context.serviceScope);
      Log.error('DocumentCentre', new Error('Error message'), this.context.serviceScope);
      Log.verbose('DocumentCentre', `VERBOSE message "${strings.BasicGroupName}"`, this.context.serviceScope); 
  
  
        <div class="${styles.welcome}">
          <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <div>${this._environmentMessage}</div>
          <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
          <div>Is ASM PowerUser: ${this.properties.asmPowerUser}</div>
          <div>Is CEN PowerUser: ${this.properties.cenPowerUser}</div>
          <div>Is CNN PowerUser: ${this.properties.cnnPowerUser}</div>
          <div>Is EMP PowerUser: ${this.properties.empPowerUser}</div>
          <div>Is HEA PowerUser: ${this.properties.heaPowerUser}</div>
          <div>Is DC PowerUser: ${this.properties.isDCPowerUser}</div>
          <div>DC URL: ${this.properties.dcURL}</div>  
          <div>Division Title: ${escape(this.properties.divisionTitle)}</div>
          <div>PP Division Selection: ${escape(this.properties.divisionName)}</div>
        </div>
  
            <div class="${ styles.description }">Slider: ${escape(this.properties.Slider)}</div>
            <div class="${ styles.description }">Toggle: ${escape(this.properties.Toggle)}</div>
            <div class="${ styles.description }">dropdowm: ${escape(this.properties.dropdowm)}</div>
            <div class="${ styles.description }">checkbox: ${escape(this.properties.checkbox)}</div>
            <div class="${ styles.description }">URL: ${escape(this.properties.URL)}</div>
            <div class="${ styles.description }">textbox: ${escape(this.properties.textbox)}</div>   
  
    }
  
    public async onInit(): Promise<void> {
      await super.onInit();
      //getSP(this.context);
      if (!Providers.globalProvider) {
        Providers.globalProvider = new SharePointProvider(this.context);
      }
      return this._getEnvironmentMessage().then((message) => {
        //this._environmentMessage = message;
      });
    }
  
    private _getEnvironmentMessage(): Promise<string> {
      if (!!this.context.sdks.microsoftTeams) {
        // running in Teams, office.com or Outlook
        return this.context.sdks.microsoftTeams.teamsJs.app
          .getContext()
          .then((context) => {
            let environmentMessage: string = "";
            switch (context.app.host.name) {
              case "Office": // running in Office
                environmentMessage = this.context.isServedFromLocalhost
                  ? strings.AppLocalEnvironmentOffice
                  : strings.AppOfficeEnvironment;
                break;
              case "Outlook": // running in Outlook
                environmentMessage = this.context.isServedFromLocalhost
                  ? strings.AppLocalEnvironmentOutlook
                  : strings.AppOutlookEnvironment;
                break;
              case "Teams": // running in Teams
                environmentMessage = this.context.isServedFromLocalhost
                  ? strings.AppLocalEnvironmentTeams
                  : strings.AppTeamsTabEnvironment;
                break;
              default:
                throw new Error("Unknown host");
            }
            return environmentMessage;
          });
      }
  
      return Promise.resolve(
        this.context.isServedFromLocalhost
          ? strings.AppLocalEnvironmentSharePoint
          : strings.AppSharePointEnvironment
      );
    }
  
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
      if (!currentTheme) {
        return;
      }
  
      //this._isDarkTheme = !!currentTheme.isInverted;
      const { semanticColors } = currentTheme;
  
      if (semanticColors) {
        this.domElement.style.setProperty(
          "--bodyText",
          semanticColors.bodyText || null
        );
        this.domElement.style.setProperty("--link", semanticColors.link || null);
        this.domElement.style.setProperty(
          "--linkHovered",
          semanticColors.linkHovered || null
        );
      }
    }
  
    protected get dataVersion(): Version {
      return Version.parse("1.0");
    }
  
    protected textBoxValidationMethod(value: string): string {
      if (value.length < 10) {
        return "Name should be at least 10 characters!";
      } else {
        return "";
      }
    }
  
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
        pages: [
          {
            //Page 1
            header: {
              description: "Page 1",
            },
            groups: [
              {
                groupName: "DC Selection",
                groupFields: [
                  PropertyPaneChoiceGroup("divisionName", {
                    label: "Division",
                    options: [
                      { key: "All", text: "All Maximus" },
                      { key: "ASM", text: "Assessments Division" },
                      { key: "CEN", text: "Central Division" },
                      { key: "CNN", text: "Connect Division" },
                      { key: "EMP", text: "Employability Division" },
                      { key: "HEA", text: "Health Division" },
                    ],
                  }),
                  
                  PropertyPaneDropdown('teamName', {
                    label:'Team',
                    options: [
                      { key: 'Item1', text: 'Item 1' },
                      { key: 'Item2', text: 'Item 2' },
                      { key: 'Item3', text: 'Item 3' }
                    ]
                  })  
                  
                ],
              },
            ],
          },
          
          { //Page 2
            header: {
              description: "Page 2"
            },
            groups: [
              {
                groupName: "Group one",
                groupFields: [
                  PropertyPaneTextField('name', {
                    label: "Name",
                    multiline: false,
                    resizable: false,
                    onGetErrorMessage: this.textBoxValidationMethod,
                    errorMessage: "This is the error message",
                    deferredValidationTime: 5000,
                    placeholder: "Please enter name","description": "Name property field"
                  }),
                  PropertyPaneTextField('description', {
                    label: "Description",
                    multiline: true,
                    resizable: true,
                    placeholder: "Please enter description","description": "Description property field"
                  })
                ]
              },
              {
                groupName: "Group two",
                groupFields: [
                  PropertyPaneSlider('Slider', {
                    label:'Slider',min:1,max:10
                  }),
                  PropertyPaneToggle('Toggle', {
                    label: 'Slider'
                  })
                ]
              },
            ]
          },
          { //Page 3
            header: {
              description: "Page 3 - URL and Label"
            },
            groups: [
              {
                groupName: "Group One",
                groupFields: [
                    PropertyPaneLink('URL',
                  { text:"Microsoft", href:'http://www.microsoft.com',target:'_blank'}),
                    PropertyPaneLabel('label',
                  { text:'Please enter designation',required:true}),
                    PropertyPaneTextField('textbox',{})
                ]
              }
            ]
          }
          
        ],
      };
    }
  }
  */