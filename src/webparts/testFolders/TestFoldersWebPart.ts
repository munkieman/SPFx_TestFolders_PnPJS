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
  divisionName: string[];
  divisionTitle: string;
  teamName: string;
  libraryName: string;
  siteArray: any;
  dcDivisions: string[]; 
  folderArray: any[];
  subFolder1Array: any[];
  subFolder2Array: any[];
  subFolder3Array: any[];
  isDCPowerUser:boolean;
}

export default class TestFoldersWebPart extends BaseClientSideWebPart<ITestFoldersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private  _getData(libraryName:string,tabNum:number,category:string): void {
    alert(libraryName);
    
    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    //let dcCount:number=0;

    this.properties.dataResults=[];
    this.properties.dcDivisions=["asm","cen"]; //,"cnn","emp","hea"];
    this.properties.divisionName=["Assessments","Central","Connect","Employability","Health"];

    const view : string = `<View>
                            <Query>
                              <Where>
                                <Or>
                                  <Eq>                
                                    <FieldRef Name="DC_Team"/>
                                    <Value Type="TaxonomyFieldType">${this.properties.siteName}</Value>
                                  </Eq>                                                                
                                  <Contains>                
                                    <FieldRef Name="DC_SharedWith"/>
                                    <Value Type='TaxonomyFieldTypeMulti'>${this.properties.siteName}</Value>
                                  </Contains>
                                </Or>                                  
                              </Where>
                              <OrderBy>
                                <FieldRef Name="DC_Division" Ascending="TRUE" />
                                <FieldRef Name="DC_Folder" Ascending="TRUE" />
                                <FieldRef Name="DC_SubFolder01" Ascending="TRUE" />
                                <FieldRef Name="DC_SubFolder02" Ascending="TRUE" />
                                <FieldRef Name="DC_SubFolder03" Ascending="TRUE" />
                                <FieldRef Name="LinkFilename" Ascending="TRUE" />
                              </OrderBy>
                            </Query>                            
                          </View>`;

    this.properties.dcDivisions.forEach(async (site,index)=>{
      //console.log(site,index);
      const dcTitle = site+"_dc";
      const webDC = Web([sp.web,`https://${this.properties.tenantURL[2]}/sites/${dcTitle}/`]); 

      await webDC.lists.getByTitle(libraryName)
        .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
        .then(async (Results) => {

          if(Results.length>0){

            console.log(dcTitle+" Results");
            console.log(Results.json());

            await this.addToResults(Results).then(async ()=>{            
              await this.getFolders(libraryName); 
            });        

          }else{
            alert("No Data found for this Team in "+dcTitle);
          }    
        })
        .catch((err) => {console.log("error",err);});
    });
  }

  private async addToResults(results:any):Promise<void>{
    let count:number=0; 
    console.log("addToResults");
    count=this.properties.dataResults.length;
    for(let x=0;x<results.length;x++){
      this.properties.dataResults[count+x]=results[x];
    }
    return;
  }

  private async getFolders(libraryName:string): Promise<void>{ //libraryName:string,tabNum:number,category:string

    const policyContainer : Element | null = this.domElement.querySelector("#policiesFolders");
    const procedureContainer : Element | null = this.domElement.querySelector("#proceduresFolders");
    const guidesContainer : Element | null = this.domElement.querySelector("#guidesFolders");
    const formsContainer : Element | null = this.domElement.querySelector("#formsFolders");
    const generalContainer : Element | null = this.domElement.querySelector("#generalFolders");
    const managementContainer : Element | null = this.domElement.querySelector("#managementFolders");
    const customContainer : Element | null = this.domElement.querySelector("#customFolders");
    
    let folderHTML: string = "";

    let folderName: string = "";
    let subFolderName1 : string = "";
    let subFolderName2 : string = "";
    let subFolderName3 : string = "";

    let folderPrev: string = "";
    let subFolderPrev1 : string = "";
    let subFolderPrev2 : string = "";
    let subFolderPrev3 : string = "";


    // *** accordion id's for folders
    let folderNameID : string = ""; 
    let subFolderName1ID: string = "";
    let subFolderName2ID: string = "";
    let subFolderName3ID: string = "";
    
    let fcount:any=0;
    let sf1count:any=0;
    let sf2count:any=0;
    let sf3count:any=0;

    // *** arrays of folder id's for the Folder EventListeners
    this.properties.folderArray = [];
    this.properties.subFolder1Array = []; 
    this.properties.subFolder2Array = [];
    this.properties.subFolder3Array = [];
    
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
        if(guidesContainer){
          guidesContainer.innerHTML="";
        }
          break;
      case "Forms":
        if(formsContainer){
          formsContainer.innerHTML="";
        }
        break;
      case "General":
        if(generalContainer){
          generalContainer.innerHTML="";
        }
        break;
      case "Management":
        if(managementContainer){
          managementContainer.innerHTML="";
        }
        break;
      case "Custom":
        if(customContainer){
          customContainer.innerHTML="";
        }
        break;
    }

    if(this.properties.dataResults.length > 0){

      for(let x=0;x<this.properties.dataResults.length;x++){
        //console.log("results item",results[x]);

        folderName = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_Folder;            
        subFolderName1 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder01;
        subFolderName2 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder02;
        subFolderName3 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder03;
      
/*
<div class="accordion accordion-flush" id="accordionFlushExample">
  <div class="accordion-item">
    <h2 class="accordion-header" id="flush-headingOne">
      <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#flush-collapseOne" aria-expanded="false" aria-controls="flush-collapseOne">
        Accordion Item #1
      </button>
    </h2>
    <div id="flush-collapseOne" class="accordion-collapse collapse" aria-labelledby="flush-headingOne" data-bs-parent="#accordionFlushExample">
      <div class="accordion-body">Placeholder content for this accordion, which is intended to demonstrate the <code>.accordion-flush</code> class. This is the first item's accordion body.</div>
    </div>
  </div>
  <div class="accordion-item">
    <h2 class="accordion-header" id="flush-headingTwo">
      <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#flush-collapseTwo" aria-expanded="false" aria-controls="flush-collapseTwo">
        Accordion Item #2
      </button>
    </h2>
    <div id="flush-collapseTwo" class="accordion-collapse collapse" aria-labelledby="flush-headingTwo" data-bs-parent="#accordionFlushExample">
      <div class="accordion-body">Placeholder content for this accordion, which is intended to demonstrate the <code>.accordion-flush</code> class. This is the second item's accordion body. Let's imagine this being filled with some actual content.</div>
    </div>
  </div>
  <div class="accordion-item">
    <h2 class="accordion-header" id="flush-headingThree">
      <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#flush-collapseThree" aria-expanded="false" aria-controls="flush-collapseThree">
        Accordion Item #3
      </button>
    </h2>
    <div id="flush-collapseThree" class="accordion-collapse collapse" aria-labelledby="flush-headingThree" data-bs-parent="#accordionFlushExample">
      <div class="accordion-body">Placeholder content for this accordion, which is intended to demonstrate the <code>.accordion-flush</code> class. This is the third item's accordion body. Nothing more exciting happening here in terms of content, but just filling up the space to make it look, at least at first glance, a bit more representative of how this would look in a real-world application.</div>
    </div>
  </div>
</div>
*/

        if(folderName!==folderPrev){

          folderHTML = '<div class="accordion accordion-flush" id="accordionFlushExample">';

          if (folderName !== '' && subFolderName1 === '' ){ 
            if(folderName.replace(/\s+/g, "")!==undefined){
              folderNameID=folderName.replace(/\s+/g, "")+"_"+x;
            }else{
              folderNameID=folderName+"_"+x;
            }
            this.properties.folderArray.push(folderName,folderNameID);

            fcount = await this.fileCount(folderName,"","","");

            folderHTML += `  <div class="accordion-item">
                              <h2 class="accordion-header" id="flush-headingOne">
                                <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#flush-collapseOne" aria-expanded="false" aria-controls="flush-collapseOne">
                                  ${folderName}
                                </button>
                              </h2>
                            </div>`;                        
          }

          if(folderName !== '' && subFolderName1 !== '' && subFolderName2 === '' && subFolderName3 === ''){ 
            sf1count = await this.fileCount(folderName,subFolderName1,"","");

            folderHTML+=`<div class="accordion-item">
                          <h2 class="accordion-header" id="folderHeader${x}">
                            <button class="btn btn-primary accordion-button" id="${folderNameID}" type="button" data-bs-toggle="collapse" data-bs-target="#accordion_${folderNameID}" aria-expanded="true" aria-controls="collapseSF1">
                              <i class="bi bi-folder2"></i><a href="#">${folderName}</a>                    
                            </button>
                          </h2>
                          <div id="accordion_${folderNameID}" class="accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordionPF${x}">
                            <div class="accordion-body">
                              <h2 class="accordion-header" id="subFolderHeader${x}">
                                <button class="btn btn-primary" id="${subFolderName1ID}" type="button">
                                  <i class="bi bi-folder2"></i>
                                  ${subFolderName1}
                                </button>
                              </h2>
                            </div>
                          </div>
                        </div>`;

          } 
          
          if (folderName !== '' && subFolderName1 !== '' && subFolderName2 !== '' && subFolderName3 === ''){ 
            sf2count = await this.fileCount(folderName,subFolderName1,subFolderName2,"");

            folderHTML+=`<div class="accordion-item">
                          <h2 class="accordion-header" id="folderHeader${x}">
                            <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#accordion_${folderNameID}" aria-expanded="true" aria-controls="collapseSF2">
                              <i class="bi bi-folder2"></i><a href="#" id="folderName">${folderName}</a>                    
                            </button>
                          </h2>
                          <div id="accordion_${folderNameID}" class="accordion-collapse collapse" aria-labelledby="folderHeader${x}" data-bs-parent="#accordionPF${x}">
                            <div class="accordion-body">
                              <div class="accordion" id="${subFolderID}">
                                <div class="accordion-item">
                                  <h2 class="accordion-header" id="subFolder1Header${x}">
                                    <button class="btn btn-primary accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#accordion_${subFolderName1ID}" aria-expanded="true" aria-controls="collapseOne">
                                      ${subFolder1Name}
                                    </button>
                                  </h2>
                                  <div id="accordion_${subFolderName1ID}" class="accordion-collapse collapse" aria-labelledby="subFolder1Header${x}" data-bs-parent="#${subFolder01ID}">
                                    <div class="accordion-body">
                                      <div class="accordion" id="accordion_${subFolderName2ID}">
                                        <div class="accordion-item">
                                          <h2 class="accordion-header" id="subFolder2Header${x}">
                                            <button class="btn btn-primary" type="button">
                                              <i class="bi bi-folder2"></i>
                                              ${subFolder2Name}
                                            </button>
                                          </h2>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>`;                                                          

          } 
          
          if (folderName !== '' && subFolderName1 !== '' && subFolderName2 !== '' && subFolderName3 !== ''){ 
            sf3count = await this.fileCount(folderName,subFolderName1,subFolderName2,subFolderName3);

            folderHTML+=`<div class="accordion-item"> 
                          <h2 class="accordion-header" id="folderHeader${x}">
                            <button class="accordion-button text-left ${styles.folderBtn}" type="button" data-bs-toggle="collapse" data-bs-target="#${folderID}" aria-expanded="true" aria-controls="documents">
                              <i class="bi bi-folder2"></i><a href="#" id="folderName" class="text-white ms-1">${folderName}</a>                    
                            </button>
                          </h2>
                          <div id="${folderID}" class="accordion-collapse collapse" aria-labelledby="folderHeader${x}" data-bs-parent="#accordionPF${x}">
                            <div class="accordion-body">
                              <div class="accordion" id="${subFolder01ID}">
                                <div class="accordion-item">
                                  <h2 class="accordion-header" id="subHeaderSF3${x}">
                                    <button class="btn btn-primary accordion-button ${styles.folderBtn}" type="button" data-bs-toggle="collapse" data-bs-target="#${subFolder02ID}" aria-expanded="true" aria-controls="collapseSF3-1">
                                      <i class="bi bi-folder2"></i><a href="#" id="subfolder1Name">${subFolder1Name}</a>                   
                                    </button>
                                  </h2>
                                <div id="collapseSF3-1" class="accordion-collapse collapse" aria-labelledby="headingSF3-1" data-bs-parent="${subFolder01ID}">
                                  <div class="accordion-body">
                                    <div class="accordion" id="${subFolder02ID}">
                                      <div class="accordion-item">
                                        <h2 class="accordion-header" id="headerSF3-2">
                                          <button class="btn btn-primary accordion-button ${styles.folderBtn}" type="button" data-bs-toggle="collapse" data-bs-target="#${subFolder03ID}" aria-expanded="true" aria-controls="collapseSF3-2">
                                            <i class="bi bi-folder2"></i><a href="#" id="subfolder2Name">${subFolder2Name}</a>                    
                                          </button>
                                        </h2>
                                        <div id="${subFolder03ID}" class="accordion-collapse collapse" aria-labelledby="subFolder3Header${x}" data-bs-parent="${subFolder02ID}">
                                          <div class="accordion-body">
                                            <h2 class="accordion-header" id="subFolder3Header${x}">
                                              <button class="btn btn-primary ${styles.folderBtn}" type="button" aria-expanded="true" aria-controls="collapseSF3-3">
                                              <i class="bi bi-folder2"></i><a href="#" id="subfolder3Name">${subFolder3Name}</a>                   
                                              </button>
                                            </h2>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>`;
          } 
          folderNamePrev=folderName;            
          
        }  
      }  // *** end of for loop

      folderHTML+=`</div></div>`;

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
          if(guidesContainer){
            guidesContainer.innerHTML=folderHTML;
          }
          break;
        case "Forms":
          if(formsContainer){
            formsContainer.innerHTML=folderHTML;
          }
          break;
        case "General":
          if(generalContainer){
            generalContainer.innerHTML=folderHTML;
          }
          break;
        case "Management":
          if(managementContainer){
            managementContainer.innerHTML=folderHTML;
          }
          break;
        case "Custom":
          if(customContainer){
            customContainer.innerHTML=folderHTML;
          }
          break;
      }
    }
  }

  private async fileCount(folder:string,subfolder1:string,subfolder2:string,subfolder3:string): Promise<number>{

    let counter : number = 0;
    let fCounter : number = 0;
    let sf1Counter : number = 0;
    let sf2Counter : number = 0;
    let sf3Counter : number = 0;

    try{
      for (let c=0;c<this.properties.dataResults.length;c++) {
        const folderName = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_Folder;            
        const subFolderName1 = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder01;
        const subFolderName2 = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder02;
        const subFolderName3 = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder03;
    
        //console.log("folder",folder,"folderName",folderName,"SubFolder1",subFolderName1);

        if ( folderName === folder && subFolderName1 === '' ){ fCounter++; }
        if ( folderName === folder && subFolderName1 === subfolder1 && subFolderName2 === '' && subFolderName3 === ''){ sf1Counter++; } 
        if ( folderName === folder && subFolderName1 === subfolder1 && subFolderName2 === subfolder2 && subFolderName3 === ''){ sf2Counter++; } 
        if ( folderName === folder && subFolderName1 === subfolder1 && subFolderName2 === subfolder2 && subFolderName3 === subfolder3){ sf3Counter++; } 
      } 
      
      if(fCounter>0){counter = fCounter;}
      if(sf1Counter>0){counter = sf1Counter;}
      if(sf2Counter>0){counter = sf2Counter;}
      if(sf3Counter>0){counter = sf3Counter;}
    } catch (err) {
      //await this.addError(this.properties.siteName, "fileCount", err.message);
    }
    return counter;
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

    console.log("tenantURL",this.properties.tenantURL);
    console.log("URL",this.properties.URL);
    console.log("siteName",this.properties.siteName);
    console.log("siteTItle",this.properties.siteTitle);
    console.log("divisionTitle",this.properties.divisionTitle);
    
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
              <div class="col-auto" id="policiesFolders"></div>
              <div class="col" id="policiesFiles"></div>
            </div>               
          </div>
          
          <div class="tab-pane fade" id="procedures" role="tabpanel" aria-labelledby="procedures" tabindex="0">
            <div class="row">
              <div class="col-auto" id="proceduresFolders"></div>
              <div class="col" id="proceduresFiles"></div>
            </div>               
          </div>

          <div class="tab-pane fade" id="guides" role="tabpanel" aria-labelledby="guidess" tabindex="0">
            <div class="row">
              <div class="col-auto" id="guidesFolders"></div>
              <div class="col" id="guidesFiles"></div>
            </div>               
          </div>

          <div class="tab-pane fade" id="forms" role="tabpanel" aria-labelledby="forms" tabindex="0">
            <div class="row">
              <div class="col-auto" id="formsFolders"></div>
              <div class="col" id="formsFiles"></div>
            </div>               
          </div>

          <div class="tab-pane fade" id="general" role="tabpanel" aria-labelledby="general" tabindex="0">
            <div class="row">
              <div class="col-auto" id="generalFolders"></div>
              <div class="col" id="generalFiles"></div>
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
