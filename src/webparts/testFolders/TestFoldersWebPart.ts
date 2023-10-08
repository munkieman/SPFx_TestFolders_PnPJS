import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
//import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './TestFoldersWebPart.module.scss';
import * as strings from 'TestFoldersWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';

import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { LogLevel, PnPLogging } from "@pnp/logging";

import * as $ from 'jquery';
require('bootstrap');
require('./styles/custom.css');

export interface ITestFoldersWebPartProps {
  description: string;
  //asmResults: any;
  //cenResults: any;
  dataResults: any;
  folderNameArray: any[];
  divisions:string[];
  URL:string;
  tenantURL: string[];
  dcURL: string;
  siteArray: string[];
  siteName: string;
  siteShortName: string;
  siteTitle: string;
  divisionName: string;
  divisionTitle: string;
  teamName: string;
  isDCPowerUser:boolean;
  folderArray: string[];
  subFolder1Array: string[];
  subFolder2Array: string[];
  subFolder3Array: string[];
  sf01IDArray: string[];
  sf02IDArray: string[];
  sf03IDArray: string[];
  filesArray : any[];
  dataFlag : boolean;

}

export default class TestFoldersWebPart extends BaseClientSideWebPart<ITestFoldersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  //private empDC = Web("https://maximusunitedkingdom.sharepoint.com/sites/emp_dc/").using(SPFx(this.context));

  private  _getData(libraryName:string,tabNum:number,category:string): void {
    alert(libraryName);
    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));    
    this.properties.dataResults=[];

    console.log(this.properties.siteTitle);

    const view : string = `<View>
                            <Query>
                              <Where>
                                <Or>
                                  <Eq>                
                                    <FieldRef Name="DC_Team"/>
                                    <Value Type="TaxonomyFieldType">${this.properties.siteTitle}</Value>
                                  </Eq>                                                                
                                  <Contains>                
                                    <FieldRef Name="DC_SharedWith"/>
                                    <Value Type='TaxonomyFieldTypeMulti'>${this.properties.siteTitle}</Value>
                                  </Contains>
                                </Or>                                  
                              </Where>
                              <OrderBy>
                                  <FieldRef Name="DC_Folder" Ascending="TRUE" />
                                  <FieldRef Name="DC_SubFolder01" Ascending="TRUE" />
                                  <FieldRef Name="DC_SubFolder02" Ascending="TRUE" />
                                  <FieldRef Name="DC_SubFolder03" Ascending="TRUE" />
                              </OrderBy>
                            </Query>                            
                          </View>`;

    this.properties.divisions=["asm","cen"]; //,"cnn","emp","hea"];

    this.properties.divisions.forEach(async (number,index)=>{
      console.log(number,index);
      const dcTitle = this.properties.divisions[index]+"_dc";
      const webDC = Web([sp.web,`https://munkieman.sharepoint.com/sites/${dcTitle}/`]); 

      await webDC.lists.getByTitle(libraryName)
        .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
        .then((Results) => {
          console.log(dcTitle+" "+Results.length);

          if(Results.length>0){
            console.log(dcTitle+" Results");
            console.log(Results);
            //this.properties.dataResults=Results;
            this._renderFolders(Results,libraryName,tabNum,category)
          }else{
            alert("No Data found for this Team in "+dcTitle);
            //const listContainer = this.domElement.querySelector('#folderContainer');
            //if(listContainer){
            //  listContainer.innerHTML = "";    
            //}
          }    
        })
        .catch(() => {});   
      });
  }
 
  private async _renderFolders(results:any,libraryName:string,tabNum:number,category:string): Promise<void>{
    alert('getting folders');
    
    let folderContainer;
    let count:any=0;
    let fCount : number = 0;

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
    //this.properties.libraryName = libraryName;
    this.properties.folderNameArray=[];
    this.properties.isDCPowerUser = true;

    console.log("Folder Results");
    console.log(results);
    console.log('libraryName=',libraryName);    

    switch (libraryName) {
      case "Policies":
        folderContainer = this.domElement.querySelector("#policiesFolders");
        break;
      case "Procedures":
        folderContainer = this.domElement.querySelector("#proceduresFolders");
        break;
      case "Guides":
        folderContainer = this.domElement.querySelector("#guidesFolders");
        break;
      case "Forms":
        folderContainer = this.domElement.querySelector("#formsFolders");
        break;
      case "General":
        folderContainer = this.domElement.querySelector("#generalFolders");
        break;
      case "Management":
        folderContainer = this.domElement.querySelector("#managementFolders");
        break;
      case "Custom":
        folderContainer = this.domElement.querySelector("#customFolders");
        break;
    }

    if(folderContainer){
      if(folderContainer.innerHTML != ""){
        folderContainer.innerHTML = "";
      }
    }

    if(results.length > 0){
      for(let x=0;x<results.length;x++){
        folderName = results[x].FieldValuesAsText.DC_x005f_Folder;
        this.properties.dataResults[x]=results[x];

        if(folderName !== null) {
          //console.log(this.properties.dataResults[x].FieldValuesAsText.DC_x005f_Folder);               
          console.log('folderName=',folderName);    
      
          if(folderName !== folderPrev){          

            // *** Parent Folder ID
            folderID = "dcTab" + tabNum + "-Folder" + fCount;

            if(folderName.replace(/\s+/g, "")==undefined){
              folderNameID=folderName;
              alert("1 "+folderNameID);
            }else{
              folderNameID=folderName.replace(/\s+/g, "");
              alert("2 "+folderNameID);
            }
        
            count=await this.fileCount(results,folderName);
            //console.log("count="+count);

            folderHTML += `<div class="accordion mt-1" id="accordionPF${x}">
                              <div class="accordion-item">
                                <h2 class="accordion-header" id="folder_${folderNameID}">
                                  <a href="" role="button" class="btn btn-primary folderBtn" id="${folderNameID}" type="button" data-bs-toggle="collapse" data-bs-target="#${folderID}" aria-expanded="true" aria-controls="accordionPF${x}">
                                    <i class="bi bi-folder2"></i>
                                    <span class="text-white ms-1">${folderName}</span>                    
                                    <span class="badge bg-secondary">${count}</span>                    
                                  </a>
                                </h2>
                                <div id="${folderID}" class="accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordionPF${x}">                                        
                                  <div class="accordion-body" id="${folderNameID}Group"></div>
                                </div>                                  
                              </div>
                            </div>`;                           
            fCount++;
            this.properties.folderNameArray.push(folderName);
            folderPrev = folderName;
          }  

          if (results[x].Knowledge_SubFolder !== null) {
            subFolderName1 = results[x].FieldValuesAsText.DC_x005f_SubFolder01;
            subFolder01ID = folderID + "-Sub01";
            this.properties.subFolder1Array.push(subFolderName1);

            if(subFolderName1 !== subFolderPrev1){  

              if(subFolderName1.replace(/\s+/g, "")==undefined){
                subFolderName1ID=subFolderName1;
              }else{
                subFolderName1ID = subFolderName1.replace(/\s+/g, "");
              }

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
              subFolderName2 = results[x].FieldValuesAsText.DC_x005f_SubFolder02;
              subFolder02ID = folderID + "-Sub02";
              this.properties.subFolder2Array.push(subFolderName2);                
              
              if(subFolderName2 !== subFolderPrev2){
                console.log("pass x="+x+" sf2="+subFolderName2+" sf2prev="+subFolderPrev2+" sfName2_ID="+subFolderName2ID);

                if(subFolderName2.replace(/\s+/g, "")==undefined){
                  subFolderName2ID=subFolderName2;
                }else{
                  subFolderName2ID = subFolderName2.replace(/\s+/g, "");
                }
  
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
                subFolderName3 = results[x].FieldValuesAsText.DC_x005f_SubFolder03;
                subFolderName3ID = subFolderName3.replace(/\s+/g, "");
                //subFolder03ID = folderID + "-Sub03";

                if(subFolderName3 !== subFolderPrev3){
                  console.log("pass x="+x+" sf3="+subFolderName3+" sf3prev="+subFolderPrev3+" sfName3_ID="+subFolderName3ID);
                  if(subFolderName2.replace(/\s+/g, "")==undefined){
                    subFolderName2ID=subFolderName2;
                  }else{
                    subFolderName2ID = subFolderName2.replace(/\s+/g, "");
                  }
  
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

    console.log("folderIDarray="+this.properties.folderNameArray);
    console.log(this.properties.dataResults);

    if(folderContainer){
      folderContainer.innerHTML = folderHTML;    
    }

    console.log("F_id=" + this.properties.sf01IDArray);
    console.log("SF1_id=" + this.properties.sf02IDArray);
    console.log("SF2_id=" + this.properties.sf03IDArray);

    if(this.properties.sf01IDArray !== undefined){
      for(let x=0;x<this.properties.sf01IDArray.length;x++){
        if(this.properties.sf01IDArray[x]!==this.properties.sf01IDArray[x-1]){
          $('#'+this.properties.sf01IDArray[x]+'Group').append(subFolder1HTML);
          const elem = document.querySelector("#"+this.properties.sf01IDArray[x]);
          elem?.classList.add('accordion-button');
        }
      }
    }
    
    if(this.properties.sf02IDArray !== undefined){
      for(let x=0;x<this.properties.sf02IDArray.length;x++){
        if(this.properties.sf02IDArray[x]!==this.properties.sf02IDArray[x-1]){
          $('#'+this.properties.sf02IDArray[x]+'Group').append(subFolder2HTML);
          const elem = document.querySelector("#"+this.properties.sf02IDArray[x]);
          elem?.classList.add('accordion-button');
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

    setTimeout(()=> {
      this.setFolderListeners();
    }
    ,3000);
  }

  private async fileCount(results:any,folderName:string): Promise<number>{

    let counter : number = 0;
    for (let c=0;c<results.length;c++) {
      if (results[c].FieldValuesAsText.DC_x005f_Folder === folderName){
        counter++;
      } 
    }
    
    //console.log(folderName +" number of files=" + counter);
    return counter;
  }

  private async setFolderListeners(): Promise<void> {
    console.log("setFolderListeners called ");
    try {
      // *** event listeners for parent folders
            
      if (this.properties.folderNameArray.length > 0) {
        for (let x = 0; x < this.properties.folderNameArray.length; x++) {         
          const folderNameID = this.properties.folderNameArray[x].replace(/\s+/g, "");
          console.log(folderNameID);

          document.getElementById("folder_" + folderNameID)
            ?.addEventListener("click", (_e: Event) => {
              alert("folder_"+folderNameID);
              //this.getFiles(folderNameID, this.properties.folderArray[x])
          });
        }
      }
    } catch (err) {
      //console.log("setFolderListeners", err.message);
      //await this.addError(this.properties.siteName, "setFolderListeners", err.message);
    }
    this.setFolderActive();
  }

  private setFolderActive(): void {

    // Get the container element
    //var btnContainer = document.getElementById("myDIV");

    // Get all buttons with class="btn" inside the container
    //let btns:any = btnContainer?.getElementsByClassName("btn");  

    // Loop through the buttons and add the active class to the current/clicked button
    //for (var i = 0; i < btns.length; i++) {
    //  btns[i].addEventListener("click", function() {
    //    var current = document.getElementsByClassName("active");

        // If there's no active class
    //    if (current.length > 0) {
    //      current[0].className = current[0].className.replace(" active", "");
    //    }

        // Add the active class to the current/clicked button
    //    this.className += " active";
    //  });
    //}
  }

  public render(): void {

    this.properties.URL = this.context.pageContext.web.absoluteUrl;
    this.properties.tenantURL = this.properties.URL.split('/',5);
    const siteSNArray : any[] = this.properties.URL.split('_',2);
    this.properties.siteShortName = siteSNArray[1];
    this.properties.siteTitle = this.context.pageContext.web.title;
    this.properties.siteArray = this.properties.siteTitle.split(" - ");
    this.properties.divisionTitle = this.properties.siteArray[0];
    this.properties.siteName = this.properties.siteArray[1];

    this.domElement.innerHTML = `
    <section class="${styles.testFolders} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
        <div>${this.properties.siteTitle}</div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
      </div>

      <div class="d-flex align-items-start">
        <div class="nav flex-column nav-pills me-3" id="v-pills-tab" role="tablist" aria-orientation="vertical">
          <button class="nav-link" id="policies_btn" data-bs-toggle="pill" data-bs-target="#policies" type="button" role="tab" aria-controls="v-pills-home" aria-selected="false">Policies</button>
          <button class="nav-link btn" id="procedures_btn" data-bs-toggle="pill" data-bs-target="#procedures" type="button" role="tab" aria-controls="v-pills-profile" aria-selected="false">Procedures</button>
          <button class="nav-link btn" id="guides_btn" data-bs-toggle="pill" data-bs-target="#guides" type="button" role="tab" aria-controls="v-pills-disabled" aria-selected="false">Guides</button>
          <button class="nav-link btn" id="forms_btn" data-bs-toggle="pill" data-bs-target="#forms" type="button" role="tab" aria-controls="v-pills-messages" aria-selected="false">Forms</button>
          <button class="nav-link btn" id="general_btn" data-bs-toggle="pill" data-bs-target="#general" type="button" role="tab" aria-controls="v-pills-settings" aria-selected="false">General</button>
        </div>

        <div class="tab-content" id="v-pills-tabContent">
          <div class="tab-pane fade libraryTab" id="policies" role="tabpanel" aria-labelledby="policies" tabindex="0"> 
            <div class="row">
              <h4 class="text-black">Folders</h4>
              <div class="justify-content-center flex-column" id="folderContainer">policies content</div>
              <div class="col-auto" id="policiesFolders"></div>
              <div class="col" id="policiesFiles"></div>
            </div>               
          </div>
          <div class="tab-pane fade libraryTab" id="procedures" role="tabpanel" aria-labelledby="procedures" tabindex="0">
            <div class="row">
              <h4 class="text-black">Folders</h4>
              <div class="justify-content-center flex-column" id="folderContainer">procedures content</div>
              <div class="col-auto" id="proceduresFolders"></div>
              <div class="col" id="proceduresFiles"></div>
            </div>               
          </div> 
        </div>
      </div>
    </section>`;

    /*
        <div class="row">
          <div class="col-6" id="docFolders">
            <h4 class="colTitle">Folder</h4>
            <div class="justify-content-center flex-column colContainer" id="folderContainer"></div>
          </div>   
          <div class="col-6" id="docFiles">
            <h4 class="colTitle">Files</h4>
            <div class="justify-content-center flex-column colContainer" id="fileContainer"></div>
          </div>
        </div>

    */
    document.getElementById('policies_btn')?.addEventListener("click", (_e:Event) => this._getData('Policies',1,""));
    document.getElementById('procedures_btn')?.addEventListener("click",(_e:Event) => this._getData('Procedures',2,""));
    document.getElementById('guides_btn')?.addEventListener("click",(_e:Event) => this._getData('Guides',3,""));
    document.getElementById('forms_btn')?.addEventListener("click",(_e:Event) => this._getData('Forms',4,""));
    document.getElementById('general_btn')?.addEventListener("click",(_e:Event) => this._getData('General',5,""));
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }
  
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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


/*

    const asmDC = Web([sp.web,`https://munkieman.sharepoint.com/sites/asm_dc/`]); 
    const cenDC = Web([sp.web,`https://munkieman.sharepoint.com/sites/cen_dc/`]); 

    asmDC.lists.getByTitle(libraryName)
      .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then((Results) => {
        //return Results.json()
        this.properties.asmResults=Results;
        console.log("ASM Results");
        console.log(Results);
      })
      .catch(() => {});

    cenDC.lists.getByTitle(libraryName)
      .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then((Results) => {
        this.properties.cenResults=Results;
        console.log("Cen Results");
        console.log(Results);
      })
      .catch(() => {});   

*/

/*
    cenDC.lists.getByTitle(libraryName)
      .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then((Results: SPHttpClientResponse) => {
        console.log(Results.json.length);

        if(Results.json.length>0){
          console.log("Central Results JSON");
          console.log(Results.json());

          //for(let x=0;x<Results.length;x++){
            this.properties.dataResults+=Results.json();
            //dataTest.push(Results[x]); 
          //}            
        }else{
          alert("No Data found for this Team in Central Document Centre");
          const listContainer = this.domElement.querySelector('#folderContainer');
          if(listContainer){
            listContainer.innerHTML = "";    
          }
        }
        console.log("CEN DC Results");
        console.log(Results);              
      });
*/    


//  private _renderFolders(response:any[]): void{

    //let dataResults:any[]=[];

    //dataResults.push(this.properties.cenResults);
    //dataResults.push(this.properties.asmResults);

    //console.log("Cen Results - Folders");
    //console.log(this.properties.cenResults);
    //console.log(dataResults[0][7]);

    
//    alert('getting folders');
    
    //let html : string = "";
    //let folderName : string = "";
    //let folderNamePrev : string = "";
    //let count = 0;
    //this.properties.folderNameIDArray=[];

//    console.log("Folder Results");
//    console.log(response);
//    console.log(response[100].FieldValuesAsText.DC_Folder);

//    for(let x=0;x<dataResults.length;x++){
//      console.log(dataResults[x]);
//      folderName = dataResults[x].FieldValuesAsText.DC_Folder;      
//      console.log('folderName='+folderName);
//    };

/*      
      if(folderName !== folderNamePrev){  
        let folderNameID=folderName.replace(/\s+/g, "")+x;
        html+=`<ul>
                <li>
                  <button class="btn btn-primary" id="${folderNameID}" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="accordionPF${x}">
                    <i class="bi bi-folder2"></i>
                    <a href="#" class="text-white ms-1">${folderName}</a>
                    <span class="badge bg-secondary">${count}</span>                    
                  </button>
                </li>
              </ul>`;            
        folderNamePrev=folderName;
        this.properties.folderNameIDArray.push(folderNameID,folderName);
        count++;
      }
      x++;
    });
    console.log("folderIDarray="+this.properties.folderNameIDArray);

    const listContainer = this.domElement.querySelector('#folderContainer');
    if(listContainer){
      listContainer.innerHTML = html;    
    }
*/
//  } 
