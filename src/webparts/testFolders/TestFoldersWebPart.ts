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

require('bootstrap');
require('./styles/custom.css');

export interface ITestFoldersWebPartProps {
  description: string;
  //asmResults: any;
  //cenResults: any;
  dataResults: any[];
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
}

export default class TestFoldersWebPart extends BaseClientSideWebPart<ITestFoldersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  //private empDC = Web("https://maximusunitedkingdom.sharepoint.com/sites/emp_dc/").using(SPFx(this.context));

  private  _getData(libraryName:string,tabNum:number,category:string): void {
    alert(libraryName);
    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));    

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

    this.properties.divisions=["asm"]; //,"cen","cnn","emp","hea"];

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
            this.properties.dataResults=Results;
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

      this._renderFolders(libraryName,tabNum,category)
  }
 
  private async _renderFolders(libraryName:string,tabNum:number,category:string): Promise<void>{
    alert('getting folders');
    
    let folderContainer;
    let fCount : number = 0;
    let folderHTML: string="";
    let folderName : string = "";
    let folderPrev : string = "";
    let folderNameID : string = "";
    let folderID : string = "";

    this.properties.folderNameArray=[];
    this.properties.isDCPowerUser = true;

    console.log("Folder Results");
    console.log(this.properties.dataResults);

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
      folderContainer.innerHTML = "";
    }

    if(this.properties.dataResults.length > 0){
      for(let x=0;x<this.properties.dataResults.length;x++){

        if(this.properties.dataResults[x].Knowledge_Folder !== null) {
          folderName = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_Folder;
          console.log(this.properties.dataResults[x].FieldValuesAsText.DC_x005f_Folder);
                
          console.log('folderName='+folderName);    
      
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
        
            let count=this.fileCount(folderName);
            console.log("count="+count);

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
        }                                    
      }
    }
    console.log("folderIDarray="+this.properties.folderNameArray);

    //const listContainer = this.domElement.querySelector('#folderContainer');
    if(folderContainer){
      folderContainer.innerHTML = folderHTML;    
    }
    setTimeout(()=> {
      this.setFolderListeners();
    }
    ,3000);
  }

  private async fileCount(folderName:string): Promise<void>{

    let counter : number = 0;
    for (let c=0;c<this.properties.dataResults.length;c++) {
      if (this.properties.dataResults[c].FieldValuesAsText.DC_x005f_Folder === folderName){
        counter++;
      } 
    }
    let html=`<span class="text-right badge bg-secondary">${counter}</span>`;
    const listContainer = this.domElement.querySelector('#folder_'+folderName.replace(/\s+/g, ""));
    if(listContainer){
      listContainer.innerHTML = html;    
    }
    
    console.log("number of files=" + counter);
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
          <button class="btn btn-primary" id="policies">Policies</button>
          <button class="btn btn-primary" id="procedures">Procedures</button>
          <button class="btn btn-primary" id="guides">Guides</button>
          <button class="btn btn-primary" id="forms">Forms</button>
          <button class="btn btn-primary" id="general">General</button>
        </div>

        <div class="tab-content" id="v-pills-tabContent">
          <div class="tab-pane fade libraryTab" id="Policies" role="tabpanel" aria-labelledby="policies" tabindex="0"> 
            <div class="row">
              <div class="col-auto" id="policiesFolders">
                <h4 class="colTitle text-black">Folder</h4>
                <div class="justify-content-center flex-column colContainer" id="folderContainer">policies content</div>
              </div>
              <div class="col" id="policiesFiles"></div>
            </div>               
          </div>
          <div class="tab-pane fade libraryTab" id="Procedures" role="tabpanel" aria-labelledby="procedures" tabindex="0">
            <div class="row">
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
    document.getElementById('policies')?.addEventListener("click", (_e:Event) => this._getData('Policies',1,""));
    document.getElementById('procedures')?.addEventListener("click",(_e:Event) => this._getData('Procedures',2,""));
    document.getElementById('guides')?.addEventListener("click",(_e:Event) => this._getData('Guides',3,""));
    document.getElementById('forms')?.addEventListener("click",(_e:Event) => this._getData('Forms',4,""));
    document.getElementById('general')?.addEventListener("click",(_e:Event) => this._getData('General',5,""));
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
