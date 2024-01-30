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
            const division = this.properties.divisionName[index];

            alert(division);

            console.log(dcTitle+" Results");
            console.log(Results.json());

            await this.addToResults(Results).then(async ()=>{            
            //const count:number = 
            //dcCount = count;
            //if(count===this.properties.dcDivisions.length){
              //console.log("count",count);
              await this._renderFolders(Results,tabNum,libraryName).then(async () => {
                  this.setFolderListeners(libraryName);
              }); //,tabNum,category,flag)
            });        
          }else{
            alert("No Data found for this Team in "+dcTitle);
          }    
        })
        .catch(() => {console.log("error")});
    });
  }

  private async addToResults(results:any):Promise<void>{
    let count:number=0; 
    //dcCount++;

    //if(results.length > 0){
      count=this.properties.dataResults.length;
      for(let x=0;x<results.length;x++){
        this.properties.dataResults[count+x]=results[x];
      }
      //console.log("results length ",results.length); 
      //console.log("dataResults length ",this.properties.dataResults.length);
      //console.log("dataResults ",this.properties.dataResults);
    //}    
    //console.log("acDivisions length ",this.properties.acDivisions.length);
    return;
  }
  
  private async _renderFolders(results:any,tabNum:number,libraryName:string): Promise<void>{ //libraryName:string,tabNum:number,category:string

    //console.log("results length ",results.length); 
    //console.log("dataResults length ",this.properties.dataResults.length);

    const policyContainer : Element | null = this.domElement.querySelector("#policiesFolders");
    const procedureContainer : Element | null = this.domElement.querySelector("#proceduresFolders");
    //const guideContainer : Element | null = this.domElement.querySelector("#guidesFolders");
    //const formContainer : Element | null = this.domElement.querySelector("#formsFolders");
    //const testContainer : Element | null = this.domElement.querySelector('#testFolders');

    //let folderCount : number = 0;
    //let sf1Count : number = 0;
    //let sf2Count : number = 0;

    //let count:number; 
    //let divisionHTML : string = "";
    let division : string = "";
    let folderHTML: string = "";

    let folderName: string = "";
    let subFolderName1 : string = "";
    let subFolderName2 : string = "";
    let subFolderName3 : string = "";

    let folderNamePrev : string = "";
    //let subFolderName1Prev : string = "";
    //let subFolderName2Prev : string = "";
    //let subFolderName3Prev : string = "";

    //let folderString : string = "";
    //let subFolder1String : string = "";
    //let subFolder2String : string = "";

    //let cen_folderPrev: string = "";
    //let asm_folderPrev: string = "";
    //let divisionPrev : string = "";

    // *** accordion id's for folders
    //let folderID : string = ""; 
    //let subFolder1ID : string = "";
    //let subFolder2ID : string = "";
    //let subFolder3ID : string = "";
    
    // *** arrays of folder id's for the Folder EventListeners
    this.properties.folderArray = [];
    this.properties.subFolder1Array = []; 
    this.properties.subFolder2Array = [];
    this.properties.subFolder3Array = [];
    
    //this.properties.libraryName = libraryName;
    //this.properties.isDCPowerUser = true;

    //if(testContainer){testContainer.innerHTML="";}

    //console.log("folder dataResults");
    //console.log(this.properties.dataResults);
    console.log("results length ",results.length);

    if(results.length > 0){

      //alert("fetching folders");

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
          this.domElement.querySelector("#guidesFolders")!.innerHTML="";
          break;
        case "Forms":
          this.domElement.querySelector("#formsFolders")!.innerHTML="";
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

      for(let x=0;x<results.length;x++){
        //console.log("results item",results[x]);

        folderName = results[x].FieldValuesAsText.DC_x005f_Folder;            
        division = results[x].FieldValuesAsText.DC_x005f_Division;

        if(folderName !== "" && division !== ""){
          subFolderName1 = results[x].FieldValuesAsText.DC_x005f_SubFolder01;
          subFolderName2 = results[x].FieldValuesAsText.DC_x005f_SubFolder02;
          subFolderName3 = results[x].FieldValuesAsText.DC_x005f_SubFolder03;

          //if(division !== divisionPrev){
          //  folderString += `<h4>${division}</h4>`;
          //  divisionPrev = division;
          //}

          if(folderName!==folderNamePrev){
            //folderString += `<div>
            //                  <h5 class="folderTitle">${folderName}</h5>
            //                </div>`;
            //this.properties.folderArray.push(results[x]);
            folderHTML += await this.makeHTML(x,division,folderName,subFolderName1,subFolderName2,subFolderName3);
            folderNamePrev=folderName;            
          }

          //if(subFolderName1!==""){
          //    folderString += `<div class="ms-1">
          //                        <h5 class="folderTitle">${subFolderName1}</h5>
          //                      </div>`;  

          //  if(subFolderName2!==""){
          //    folderString += `<div class="ms-2">
          //                        <h5 class="folderTitle">${subFolderName2}</h5>
          //                      </div>`;

          //    if(subFolderName3 !==""){                
          //        folderString += `<div class="ms-2">
          //                            <h5 class="folderTitle">${subFolderName2}</h5>
          //                          </div>`;
          //    }
          //  }
          //}
    
          //if(testContainer){testContainer.innerHTML=folderString;}     
          
/*          
          switch(division){
            case "Assessments":
              
              // *** check is a new folder, if so create new folder string and add to DOM
              if(folderName !== asm_folderPrev){

                if(division !== divisionPrev){
                  folderHTML+=`<h4>${division}</h4>`;
                } 

                folderHTML += await this.makeHTML(x,division,folderName,subFolderName1,subFolderName2,subFolderName3);
                asm_folderPrev = folderName;
              }                          
              divisionPrev = division;
              break;

            case "Central":

              // *** check is a new folder, if so create new folder string and add to DOM
              if(folderName !== cen_folderPrev){
                
                if(division !== divisionPrev){
                  folderHTML+=`<h4>${division}</h4>`;
                }

                folderHTML += await this.makeHTML(x,division,folderName,subFolderName1,subFolderName2,subFolderName3);               
                cen_folderPrev = folderName;
              }
              divisionPrev = division;
              break;
          }
*/         
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
          this.domElement.querySelector("#guidesFolders")!.innerHTML=folderHTML;
          break;
        case "Forms":
          this.domElement.querySelector("#formsFolders")!.innerHTML=folderHTML;
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

/*          
          if(division !== divisionPrev){
            if(folderName !== cen_folderPrev || folderName !== asm_folderPrev){

              if(division !== divisionPrev){
                folderHTML+=`<h4>${division}</h4>`;
              }

              if(subFolderName1 !== ""){
                console.log(division," ",folderName," ",subFolderName1);
                folderHTML+=`<a href="#" class="text-white ms-1" >${folderName}</a>`;
              }else{
                console.log(division," ",folderName);
                folderHTML+=`<a href="#" class="text-white ms-1" >${folderName}</a>`;
              }
            }

            switch(division){
              case 'Assessments':
                asm_folderPrev = folderName;
                break;
              case 'Central':
                cen_folderPrev = folderName;
                break;
            }
          }
          divisionPrev = division;  
          
          // *** check is a new folder, if so create new folder string and add to DOM
          if(folderName !== folderPrev){
              
            if(division !== divisionPrev){
              folderHTML+=`<h4>${division}</h4>`;
            }
            
            // *** check if folderName has spaces or special characters and remove them for the ID.
            if(folderName.replace(/\s+/g, "")!==undefined){
              folderNameID=folderName.replace(/\s+/g, "")+"_"+x;
            }else{
              folderNameID=folderName+"_"+x;
            }
            this.properties.folderArray.push(folderName,folderNameID);
            fcount = await this.fileCount(folderName);

            //console.log("CHK folder ",folderName);          
*/
 
  private async makeHTML(x:number,division:string,folderName:string,subFolderName1:string,subFolderName2:string,subFolderName3:string):Promise<string>{

    //alert("make html "+division);
    
    let fcount:any=0;
    let sf1count:any=0;
    let sf2count:any=0;
    let sf3count:any=0;

    let folderHTML: string = "";
    let folderHTMLEnd : string = "";
    
    // *** folder id's for event listeners on button click
    let folderNameID : string = ""; 
    let subFolderName1ID: string = "";
    let subFolderName2ID: string = "";
    let subFolderName3ID: string = "";

    // *** check if folderName has spaces or special characters and remove them for the ID.
    if(folderName.replace(/\s+/g, "")!==undefined){
      folderNameID=folderName.replace(/\s+/g, "")+"_"+x;
    }else{
      folderNameID=folderName+"_"+x;
    }
    this.properties.folderArray.push(division,folderName,folderNameID);
    fcount = await this.fileCount(division,folderName);

    if(subFolderName1!==``){       
      folderHTML+=`<div class="accordion" id="accordionPF-${x}">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="folder_${folderNameID}">
                        <button class="btn btn-primary accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF1-${x}" aria-expanded="true" aria-controls="collapseSF1-${x}">
                          <i class="bi bi-folder2"></i>
                          <a href="#" class="text-white ms-1" id="${folderNameID}">${folderName}</a>
                          <span class="badge bg-secondary">${fcount}</span>                    
                        </button>
                      </h2>`;
    }else{
      folderHTML+=`<div class="accordion" id="accordionPF-${x}">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="folder_${folderNameID}">
                        <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseSF1-${x}">
                          <i class="bi bi-folder2"></i>
                          <a href="#" class="text-white ms-1" id="${folderNameID}">${folderName}</a>
                          <span class="badge bg-secondary">${fcount}</span>                    
                        </button>
                      </h2>`;
    }
    
    if(subFolderName1 !== ''){
      //console.log("CHK subfolder1 ",subFolderName1);

      // *** check if subfolderName has spaces or special characters and remove them for the ID.
      if(subFolderName1.replace(/\s+/g, "")!==undefined){
        subFolderName1ID=subFolderName1.replace(/\s+/g, "")+"_"+x;
      }else{
        subFolderName1ID=subFolderName1+"_"+x;
      }
      this.properties.subFolder1Array.push(division,subFolderName1,subFolderName1ID);
      sf1count = await this.fileCount(division,subFolderName1);

      if(subFolderName2 !== ``){
        folderHTML+=`<div id="collapseSF1-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF1-${x}" data-bs-parent="#accordionPF-${x}">
                      <div class="accordion-body"> 
                        <div class="accordion" id="accordionSF1-${x}">                              
                          <div class="accordion-item">
                            <h2 class="accordion-header" id="SF1_${subFolderName1ID}">
                              <button class="btn btn-primary accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF2-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                <i class="bi bi-folder2"></i>
                                <a href="#" class="text-white ms-1" id="${subFolderName1ID}">${subFolderName1}</a>
                                <span class="badge bg-secondary">${sf1count}</span>                                        
                              </button>
                            </h2>`;
        folderHTMLEnd+=`</div></div></div></div>`;

      }else{
        folderHTML+=`<div id="collapseSF1-${x}" class="ms-1 accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordionPF-${x}">
                      <div class="accordion-body">
                        <div class="accordion-item">
                          <h2 class="accordion-header" id="SF1_${subFolderName1ID}">
                            <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseSF1">
                              <i class="bi bi-folder2"></i>
                              <a href="#" class="text-white ms-1" id="${subFolderName1ID}">${subFolderName1}</a>
                              <span class="badge bg-secondary">${sf1count}</span>                    
                            </button>
                          </h2>
                        </div>
                      </div>
                    </div>`;
      }               
    }

    if(subFolderName2 !== ''){
      //console.log("CHK subfolder2 ",subFolderName2);  

      // *** check if subfolderName has spaces or special characters and remove them for the ID.
      if(subFolderName2.replace(/\s+/g, "")!==undefined){
        subFolderName2ID=subFolderName2.replace(/\s+/g, "")+"_"+x;
      }else{
        subFolderName2ID=subFolderName2+"_"+x;
      }
      this.properties.subFolder2Array.push(division,subFolderName2,subFolderName2ID);
      sf2count = await this.fileCount(division,subFolderName2);

      if(subFolderName3 !==``){
        folderHTML+=`<div id="collapseSF2-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF2" data-bs-parent="accordionSF1-${x}">
                      <div class="accordion-body">
                        <div class="accordion" id="accordionSF2-${x}">
                          <div class="accordion-item">
                            <h2 class="accordion-header" id="SF2_${subFolderName2ID}">
                              <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF3-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                <i class="bi bi-folder2"></i>
                                <a href="#" class="text-white ms-1">${subFolderName2}</a>
                                <span class="badge bg-secondary">${sf2count}</span>                    
                              </button>
                            </h2>`;
        folderHTMLEnd+=`</div></div></div></div>`;

      }else{
        folderHTML+=`<div id="collapseSF2-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF2" data-bs-parent="accordionSF1-${x}">
                      <div class="accordion-body">
                        <div class="accordion-item">
                          <h2 class="accordion-header" id="SF2_${subFolderName2ID}">
                            <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="false" aria-controls="collapseSF2-${x}">
                              <i class="bi bi-folder2"></i>
                              <a href="#" class="text-white ms-1">${subFolderName2}</a>
                              <span class="badge bg-secondary">${sf2count}</span>                    
                            </button>
                          </h2>
                        </div>
                      </div>
                    </div>`;
      }               
    }   

    if(subFolderName3 !== ''){

      // *** check if subfolderName has spaces or special characters and remove them for the ID.
      if(subFolderName3.replace(/\s+/g, "")!==undefined){
        subFolderName3ID=subFolderName3.replace(/\s+/g, "")+"_"+x;
      }else{
        subFolderName3ID=subFolderName3+"_"+x;
      }
      this.properties.subFolder3Array.push(division,subFolderName3,subFolderName3ID);
      sf3count = await this.fileCount(division,subFolderName3);

      folderHTML+=`<div id="collapseSF3-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF3" data-bs-parent="accordionSF2-${x}">
                    <div class="accordion-body">
                      <div class="accordion-item">
                        <h2 class="accordion-header" id="SF3_${subFolderName3ID}">
                          <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="false" aria-controls="collapseSF3-${x}">
                            <i class="bi bi-folder2"></i>
                            <a href="#" class="text-white ms-1" id="sf3ID">${subFolderName3}}</a>
                            <span class="badge bg-secondary">${sf3count}</span>                    
                          </button>
                        </h2>
                      </div>
                    </div>
                  </div>`;
    }
    folderHTML+=folderHTMLEnd;
    return folderHTML;
  }

  private async fileCount(division:string,folderName:string): Promise<number>{

    let counter : number = 0;
    for (let c=0;c<this.properties.dataResults.length;c++) {
      if(this.properties.dataResults[c].FieldValuesAsText.DC_x005f_Division === division){
        if (this.properties.dataResults[c].FieldValuesAsText.DC_x005f_Folder === folderName){
          counter++;
        } 
        if (this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder01 === folderName){
          counter++;
        } 
        if (this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder02 === folderName){
          counter++;
        } 
        if (this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder03 === folderName){
          counter++;
        }
      } 
    }
    return counter;
  } 

  private setFolderListeners(libraryName:string): void {
    //console.log("setFolderListeners called");
    
    //this.properties.getFilesCallFlag = false;
    //console.log("setfolderlisteners callflag="+this.properties.getFilesCallFlag);

    try {

      // *** event listeners for parent folders      
      if (this.properties.folderArray.length > 0) {

        //const folderNameID = this.properties.folderArray.filter(function (value, index, ar) {
        //  console.log("folderNameID ",index % 3);
        //  return (index % 3 > 0);
        //});

        //const folderName = this.properties.folderArray.filter(function (value, index, ar) {         
        //  return (index % 3 === 0);
        //});

        //const division = this.properties.folderArray.filter(function (value, index, ar) {
        //  if(index % 3 === 0){
        //    console.log("division ",value);
        //  }
        //  return (index % 3 === 0);
        //});

        const folderNameID = this.properties.folderArray.filter((value, index) => index % 3 === 3 - 1);
        const folderName = this.properties.folderArray.filter((value, index) => index % 3 === 3 - 2);
        //const division = this.properties.folderArray.filter((value, index) => index % 3 === 3 - 3);

        console.log("folderNameID ",folderNameID);
        console.log("folderName ",folderName);
        //console.log("division ",division);

        for (let x = 0; x < folderNameID.length; x++) {
          document.getElementById("folder_" + folderNameID[x])
            ?.addEventListener("click", (_e: Event) => {
              this.getFiles(libraryName,folderName[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#folder_" + folderNameID[x]);
              elem?.classList.add('active');
            });
        }
      }

      // *** event listeners for subfolders level 1      
      if (this.properties.subFolder1Array.length > 0) {
        const subFolder1NameID = this.properties.subFolder1Array.filter((value, index) => index % 3 === 3 - 1);
        const subFolder1Name = this.properties.subFolder1Array.filter((value, index) => index % 3 === 3 - 2);
        //const subFolder1division = this.properties.folderArray.filter((value, index) => index % 3 === 3 - 3);

        for (let x = 0; x < subFolder1NameID.length; x++) {
          document.getElementById("SF1_" + subFolder1NameID[x])
            ?.addEventListener("click", (_e: Event) => {
              //console.log("subfolder1name ",subFolder1Name[x]);
              this.getFiles(libraryName,subFolder1Name[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#SF1_" + subFolder1NameID[x]);
              elem?.classList.add('active');

            });
        }
      }

      // *** event listeners for subfolders level 2      
      if (this.properties.subFolder2Array.length > 0) {

        const subFolder2NameID = this.properties.subFolder2Array.filter(function (value, index, ar) {
          return (index % 2 > 0);
        });

        const subFolder2Name = this.properties.subFolder2Array.filter(function (value, index, ar) {
          return (index % 2 === 0);
        });

        for (let x = 0; x < subFolder2NameID.length; x++) {
          
          document.getElementById("SF2_" + subFolder2NameID[x])
            ?.addEventListener("click", (_e: Event) => {
              this.getFiles(libraryName,subFolder2Name[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#SF2_" + subFolder2NameID[x]);
              elem?.classList.add('active');

            });
        }
      }      

      // *** event listeners for subfolders level 3      
      if (this.properties.subFolder3Array.length > 0) {

        const subFolder3NameID = this.properties.subFolder3Array.filter(function (value, index, ar) {
          return (index % 2 > 0);
        });

        const subFolder3Name = this.properties.subFolder3Array.filter(function (value, index, ar) {
          return (index % 2 === 0);
        });

        for (let x = 0; x < subFolder3NameID.length; x++) {
          
          document.getElementById("SF3_" + subFolder3NameID[x])
            ?.addEventListener("click", (_e: Event) => {
              this.getFiles(libraryName,subFolder3Name[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#SF3_" + subFolder3NameID[x]);
              elem?.classList.add('active');

            });
        }
      } 

    } catch (err) {
      //console.log("setFolderListeners", err.message);
      //this.addError(this.properties.siteName, "setFolderListeners", err.message);
    }    
    //console.log("setFolderListeners completed");
  }   

  private getFiles(libraryName:string,folderName:string) :void {
    //alert("getfiles for "+folderName);
    
    let divisionPrev : string = "";
    let fileHTML: string = "";
    let powerUserTools: string = "";
    let draftText: string = "";
    let fileTypeIcon: string = "";

    const policyContainer : Element | null = this.domElement.querySelector("#policiesFiles");
    const procedureContainer : Element | null = this.domElement.querySelector("#proceduresFiles");

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

        break;
      case "Forms":
  
        break;
      case "General":

        break;
      case "Management":

        break;
      case "Custom":

        break;
    }

    for(let f=0;f<this.properties.dataResults.length;f++){
      let fileFlag : boolean = false;  
      const division : string = this.properties.dataResults[f].FieldValuesAsText.DC_x005f_Division;
      const Folder : string = this.properties.dataResults[f].FieldValuesAsText.DC_x005f_Folder;
      const SubFolder1 : string = this.properties.dataResults[f].FieldValuesAsText.DC_x005f_SubFolder01;
      const SubFolder2 : string = this.properties.dataResults[f].FieldValuesAsText.DC_x005f_SubFolder02;
      const SubFolder3 : string = this.properties.dataResults[f].FieldValuesAsText.DC_x005f_SubFolder03;
      const fileName : string = this.properties.dataResults[f].FieldValuesAsText.FileLeafRef;
      const fileStatus: string = this.properties.dataResults[f].FieldValuesAsText.OData__x005f_ModerationStatus;
      const fileType: string = this.properties.dataResults[f].FieldValuesAsText.File_x005f_x0020_x005f_Type;
      const fileURL: string = this.properties.dataResults[f].ServerRedirectedEmbedUrl;
      const fileID: string = this.properties.dataResults[f].ID;

      if(Folder === folderName && SubFolder1 === ''){
        console.log("folderName ",folderName," Folder ",Folder);
        fileFlag = true;
      }

      if( SubFolder1 === folderName && SubFolder2 === '' ){
        console.log("subFolder ",folderName," SubFolder1 ",SubFolder1);
        fileFlag = true;
      }

      if( SubFolder2 === folderName && SubFolder3 === "" ){
        fileFlag = true;
      }

      if( SubFolder3 === folderName){
        fileFlag = true;
      }
      
      if(fileFlag){
  
        if(division !== divisionPrev){
          fileHTML+=`<h4>${division}</h4>`;
        }        
        divisionPrev = division;

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

        switch (libraryName) {
          case "Policies":
            if(policyContainer){
              policyContainer.innerHTML=fileHTML;
            }
            break;
          case "Procedures":
            if(procedureContainer){
              procedureContainer.innerHTML=fileHTML;
            }
            break;
          case "Guides":

            break;
          case "Forms":

            break;
          case "General":

            break;
          case "Management":

            break;
          case "Custom":

            break;
        }                             
      }
    }
    return;
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

      <div>
        <h4>Test Folders</h4>
        <div id="testFolders"></div>
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
