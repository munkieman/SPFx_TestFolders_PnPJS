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

//import * as $ from 'jquery';
require('bootstrap');
require('./styles/custom.css');

export interface ITestFoldersWebPartProps {
  description: string;
  dataResults: any[];
  siteTitle: string;
  completeFlag: boolean;
  URL:string;
  tenantURL: any;
  dcURL: string;   
  siteName: string;
  siteShortName: string;
  divisionName: string;
  divisionTitle: string;
  teamName: string;
  libraryName: string;
  siteArray: any;
  acDivisions: string[]; 
}

export default class TestFoldersWebPart extends BaseClientSideWebPart<ITestFoldersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private  _getData(libraryName:string,tabNum:number,category:string): void {
    alert(libraryName);
    
    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    let dcCount:number=0;

    this.properties.dataResults=[];
    this.properties.acDivisions=["asm","cen"]; //,"cnn","emp","hea"];

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
                                  <FieldRef Name="LinkFilename" Ascending="TRUE" />
                              </OrderBy>
                            </Query>                            
                          </View>`;

    this.properties.acDivisions.forEach(async (site,index)=>{
      console.log(site,index);
      const dcTitle = site+"_dc";
      const webDC = Web([sp.web,`https://munkieman.sharepoint.com/sites/${dcTitle}/`]); 

      await webDC.lists.getByTitle(libraryName)
        .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
        .then(async (Results) => {

          if(Results.length>0){
            //console.log(dcTitle+" Results");
            //console.log(dcTitle+" "+Results.length);
            const count:number = await this.addToResults(Results,dcCount); //.then(async ()=>{            
            dcCount = count;
            if(count===this.properties.acDivisions.length){
              //console.log("count",count);
              await this._renderFolders(Results,libraryName); //,tabNum,category,flag)
            }        
          }else{
            alert("No Data found for this Team in "+dcTitle);
          }    
        })
        .catch(() => {console.log("error")});
    });
  }

  private async addToResults(results:any,dcCount:number):Promise<number>{
    let count:number=0; 
    dcCount++;

    if(results.length > 0){
      count=this.properties.dataResults.length;
      for(let x=0;x<results.length;x++){
        this.properties.dataResults[count+x]=results[x];
      }
      //console.log("results length ",results.length); 
      //console.log("dataResults length ",this.properties.dataResults.length);
      console.log("dataResults ",this.properties.dataResults);
    }    
    //console.log("acDivisions length ",this.properties.acDivisions.length);
    return dcCount;
  }
  
  private async _renderFolders(results:any,libraryName:string): Promise<void>{ //libraryName:string,tabNum:number,category:string

    console.log("results length ",results.length); 
    console.log("dataResults length ",this.properties.dataResults.length);

    const policyContainer : Element | null = this.domElement.querySelector("#policiesFolders");
    const procedureContainer : Element | null = this.domElement.querySelector("#proceduresFolders");

    //let fileCount:any=0;
    //let folderCount : number = 0;
    //let count:number; 

    //let HTML : string = "<div class='row'>";
    let folderHTML: string = "";
    let folderHTMLEnd : string = "";

    let folderName: string = "";
    let subFolderName1 : string = "";
    let subFolderName2 : string = "";
    let subFolderName3 : string = "";

    let folderPrev: string = "";
    //let subFolderPrev1: string = "";
    //let subFolderPrev2: string = "";
    //let subFolderPrev3: string = "";

    // *** folder id's for event listeners on button click
    //let folderNameID : string = ""; 
    //let subFolderName1ID: string = "";
    //let subFolderName2ID: string = "";
    //let subFolderName3ID: string = "";

    // *** accordion id's for folders
    //let folderID : string = ""; 
    //let subFolder01ID : string = "";
    //let subFolder02ID : string = "";
    //let subFolder03ID : string = "";

   //this.properties.dataResults = [];
 
    // *** used for the Folder EventListeners
    //this.properties.folderArray = [];
    //this.properties.subFolder1Array = []; 
    //this.properties.subFolder2Array = [];
    //this.properties.subFolder3Array = [];

    // *** used for the Folder HTML code
    //this.properties.folderNameArray=[];
    //this.properties.sf01IDArray=[]; 
    //this.properties.sf02IDArray=[];
    //this.properties.sf03IDArray=[];
    
    //this.properties.libraryName = libraryName;
    //this.properties.isDCPowerUser = true;

//    if(results.length > 0){
//      count=this.properties.dataResults.length;
//      for(let x=0;x<results.length;x++){
//        this.properties.dataResults[count+x]=results[x];
//      }
//      console.log("count ",count); 
//      console.log("results length ",results.length); 
//      console.log("dataResults length ",this.properties.dataResults.length);
//      console.log("Flag ",flag);
//      if(flag===false){return};
//    }

    console.log("folder dataResults");
    console.log(this.properties.dataResults);
    
    if(this.properties.dataResults.length > 0){

      alert("fetching folders");

      switch (libraryName) {
        case "Policies":        
          if(policyContainer){
            policyContainer.innerHTML="";
          }
          break;
        case "Procedures":
          if(procedureContainer){
            procedureContainer.innerHTML="";
          }
          break;
        case "Guides":
          //this.domElement.querySelector("#guidesFolders")!.innerHTML="";
          break;
        case "Forms":
          //this.domElement.querySelector("#formsFolders")!.innerHTML="";
          break;
        case "General":
          //this.domElement.querySelector("#generalFolders")!.innerHTML="";
          break;
        case "Management":
          //this.domElement.querySelector("#managementFolders")!.innerHTML="";
          break;
        case "Custom":
          //this.domElement.querySelector("#customFolders")!.innerHTML="";
          break;
      }

      //if(this.domElement.querySelector("#policiesFolders")!=null) {
      //  this.domElement.querySelector("#policiesFolders")!.innerHTML = "<h2>folder name</h2>";
      //}

      for(let x=0;x<this.properties.dataResults.length;x++){
        folderName = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_Folder;            

        if(folderName !== ""){
          subFolderName1 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder01;
          subFolderName2 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder02;
          subFolderName3 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder03;
        
          // *** check is a new folder, if so create new folder string and add to DOM
          if(folderName !== folderPrev){

            console.log("CHK folder ",folderName);
            if(subFolderName1!==``){       
              folderHTML+=`<div class="accordion" id="accordion${x}">
                            <div class="accordion-item">
                              <h2 class="accordion-header" id="headerPF-${x}">
                                <button class="btn btn-primary accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF1-${x}" aria-expanded="true" aria-controls="collapseSF1-${x}">
                                  <i class="bi bi-folder2"></i>
                                  <a href="#" class="text-white ms-1">${folderName}</a>
                                  <span class="badge bg-secondary">0</span>                    
                                </button>
                              </h2>`;
            }else{
              folderHTML+=`<div class="accordion" id="accordion${x}">
                            <div class="accordion-item">
                              <h2 class="accordion-header" id="headerPF-${x}">
                                <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseOne">
                                  <i class="bi bi-folder2"></i>
                                  <a href="#" class="text-white ms-1">${folderName}</a>
                                  <span class="badge bg-secondary">0</span>                    
                                </button>
                              </h2>`;
            }
            folderPrev = folderName;

            if(subFolderName1 !== ''){
              console.log("CHK subfolder1 ",subFolderName1);

              if(subFolderName2 !== ``){
                folderHTML+=`<div id="collapseSF1-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordion${x}">
                              <div class="accordion-body"> 
                                <div class="accordion" id="accordionSF1-${x}">                              
                                  <div class="accordion-item">
                                    <h2 class="accordion-header" id="headerSF1-${x}">
                                      <button class="btn btn-primary accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF2-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${subFolderName1}</a>
                                        <span class="badge bg-secondary">0</span>                                        
                                      </button>
                                    </h2>`;
                folderHTMLEnd+=`</div></div></div></div>`;

              }else{
                folderHTML+=`<div id="collapseSF1-${x}" class="ms-1 accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordion${x}">
                              <div class="accordion-body">
                                <div class="accordion" id="accordionSF1-${x}">
                                  <div class="accordion-item">
                                    <h2 class="accordion-header" id="headerSF1-${x}">
                                      <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseSF1">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${subFolderName1}</a>
                                        <span class="badge bg-secondary">0</span>                    
                                      </button>
                                    </h2>
                                  </div>
                                </div>
                              </div>
                            </div>`;
              }               
            }

            if(subFolderName2 !== ''){
              console.log("CHK subfolder2 ",subFolderName2);  

              if(subFolderName3 !==``){
                folderHTML+=`<div id="collapseSF2-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF2" data-bs-parent="accordionSF1-${x}">
                              <div class="accordion-body">
                                <div class="accordion" id="accordionSF2-${x}">
                                  <div class="accordion-item">
                                    <h2 class="accordion-header" id="headerSF2-${x}">
                                      <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF2-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${subFolderName2}</a>
                                        <span class="badge bg-secondary">0</span>                    
                                      </button>
                                    </h2>`;
                folderHTMLEnd+=`</div></div></div></div>`;

              }else{
                folderHTML+=`<div id="collapseSF2-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF2" data-bs-parent="accordionSF1-${x}">
                              <div class="accordion-body">
                                <div class="accordion" id="accordionSF2-${x}">
                                  <div class="accordion-item">
                                    <h2 class="accordion-header" id="headerSF2-${x}">
                                      <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${subFolderName2}</a>
                                        <span class="badge bg-secondary">0</span>                    
                                      </button>
                                    </h2>
                                  </div>
                                </div>
                              </div>
                            </div>`;
              }               
            }   

            if(subFolderName3 !== ''){
              folderHTML+=`<div id="collapseSF3-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF3" data-bs-parent="accordionSF2-${x}">
                            <div class="accordion-body">
                              <h2 class="accordion-header" id="headerSF3-${x}">
                                <button 
                                  class="btn btn-primary" 
                                  type="button" 
                                  data-bs-toggle="collapse" 
                                  data-bs-target="#collapseSF3-${x}" 
                                  aria-expanded="false" 
                                  aria-controls="collapseSF3-${x}">
                                    <i class="bi bi-folder2"></i>
                                    <a href="#" id="sf3ID"> 
                                      ${subFolderName3}}
                                    </a>
                                </button>
                              </h2>
                            </div>
                          </div>`;
            }
            folderHTML+=folderHTMLEnd;
          }
        }          
      }  // *** end of for loop

      folderHTML+=`</div></div>`;
      console.log(folderHTML);

      switch (libraryName) {
        case "Policies":
          if(policyContainer){
            policyContainer.innerHTML=folderHTML;
          }
          break;
        case "Procedures":
          if(procedureContainer){
            procedureContainer.innerHTML=folderHTML;
          }
          break;
        case "Guides":
          //this.domElement.querySelector("#guidesFolders")!.innerHTML=folderHTML;
          break;
        case "Forms":
          //this.domElement.querySelector("#formsFolders")!.innerHTML=folderHTML;
          break;
        case "General":
          //this.domElement.querySelector("#generalFolders")!.innerHTML=folderHTML;
          break;
        case "Management":
          //this.domElement.querySelector("#managementFolders")!.innerHTML=folderHTML;
          break;
        case "Custom":
          //this.domElement.querySelector("#customFolders")!.innerHTML=folderHTML;
          break;
      }
    }
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
    this.properties.completeFlag = false;

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
          <div class="tab-pane fade" id="policies" role="tabpanel" aria-labelledby="policies" tabindex="0"> 
            <div class="row">
              <div class="col-auto" id="policiesFolders">
                <h4 class="text-black">Folders</h4>
              </div>
              <div class="col" id="policiesFiles">
                <h4 class="text-black">Files</h4>
              </div>
            </div>               
          </div>
          <div class="tab-pane fade" id="procedures" role="tabpanel" aria-labelledby="procedures" tabindex="0">
            <div class="row">
              <h4 class="text-black">Folders</h4>
              <div class="col-auto" id="proceduresFolders"></div>
              <div class="col" id="proceduresFiles"></div>
            </div>               
          </div> 
        </div>
      </div>
    </section>`;

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
                        folderHTML+=`<div class="row">
                                      <button class="btn btn-primary" type="button" data-bs-toggle="collapse">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${folderName}</a>
                                        <span class="badge bg-secondary">0</span>                    
                                      </button>
                                    </div>
                                    <div class="row ms-1">
                                      <button class="btn btn-primary" type="button" data-bs-toggle="collapse">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${subFolderName1}</a>
                                        <span class="badge bg-secondary">0</span>                    
                                      </button>
                                    </div>                
                                    <div class="row ms-2">
                                      <button class="btn btn-primary" type="button" data-bs-toggle="collapse">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${subFolderName2}</a>
                                        <span class="badge bg-secondary">0</span>                    
                                      </button>
                                    </div>                
                                    <div class="row ms-3">
                                      <button class="btn btn-primary" type="button" data-bs-toggle="collapse">
                                        <i class="bi bi-folder2"></i>
                                        <a href="#" class="text-white ms-1">${subFolderName3}</a>
                                        <span class="badge bg-secondary">0</span>                    
                                      </button>
                                    </div>`;                
*/
               
/*
        if(folderName !== "" && folderName !== folderPrev){
          if(subFolderName1 !== "" && subFolderName1 !== subFolderPrev1){
            if(subFolderName2 !== "" && subFolderName2 !== subFolderPrev2){
              if(subFolderName3 !== "" && subFolderName3 !== subFolderPrev3){
                console.log("CHK folder ",folderName);
                console.log("CHK subfolder1 ",subFolderName1);
                console.log("CHK subfolder2 ",subFolderName2); 
                console.log("CHK subfolder3 ",subFolderName3);

                //folderPrev = folderName;              
                //subFolderPrev1 = subFolderName1;
                //subFolderPrev2 = subFolderName2;
                //subFolderPrev3 = subFolderName3;      
              }else{
                console.log("CHK folder ",folderName);
                console.log("CHK subfolder1 ",subFolderName1);
                console.log("CHK subfolder2 ",subFolderName2); 
                //folderPrev = folderName;              
                //subFolderPrev1 = subFolderName1;
                //subFolderPrev2 = subFolderName2;
              } // *** end subfolder3 check
            }else{ 
              console.log("CHK folder ",folderName);
              console.log("CHK subfolder1 ",subFolderName1);
              //folderPrev = folderName;              
              //subFolderPrev1 = subFolderName1;
            } // *** end subfolder2 check                        
          } else {  
            //folderHTML+=`<div class="row">
            //              <button class="btn btn-primary" type="button" data-bs-toggle="collapse">
            //                <i class="bi bi-folder2"></i>
            //                <a href="#" class="text-white ms-1">${folderName}</a>
            //                <span class="badge bg-secondary">0</span>                    
            //              </button>
            //            </div>`;                
            console.log("CHK folder ",folderName);
            //folderPrev = folderName;
          }// *** end subfolder1 check                           
        } // *** end folder check

*/
