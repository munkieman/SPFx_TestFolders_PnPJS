/* 
 //alert('getting folders');
    
    //let folderContainer;
    //let count:any=0;
    //let fCount : number = 0;

    //let folderHTML: string = "";
    //let subFolder1HTML : string = "";
    //let subFolder2HTML : string = "";
    //let subFolder3HTML : string = "";

    //let folderName: string = "";
    //let subFolderName1 : string = "";
    //let subFolderName2 : string = "";
    //let subFolderName3 : string = "";

    //let folderPrev: string = "";
    //let subFolderPrev1: string = "";
    //let subFolderPrev2: string = "";
    //let subFolderPrev3: string = "";

    //let folderNameID : string = ""; // folder id's for event listeners on button click
    //let subFolderName1ID: string = "";
    //let subFolderName2ID: string = "";
    //let subFolderName3ID: string = "";
    
    //let folderID : string = ""; // accordion id's for folders
    //let subFolder01ID : string = "";
    //let subFolder02ID : string = "";
    //let subFolder03ID : string = "";

    //this.properties.dataResults = [];
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

    //console.log("Folder Results");
    //console.log(results);
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

    //if(folderContainer){
    //  if(folderContainer.innerHTML !== ""){
    //    folderContainer.innerHTML = "";
    //  }
    //}

//    if(results.length > 0){
//      for(let x=0;x<results.length;x++){

//        folderName = results[x].FieldValuesAsText.DC_x005f_Folder;
//        subFolderName1 = results[x].FieldValuesAsText.DC_x005f_SubFolder01;
//        subFolderName2 = results[x].FieldValuesAsText.DC_x005f_SubFolder02;
//        subFolderName3 = results[x].FieldValuesAsText.DC_x005f_SubFolder03;

//        this.properties.dataResults[x]=results[x];

//        if(folderName !== undefined && subFolderName1 === undefined){
//          if(folderName !== folderPrev){
//            console.log("FolderName ",folderName);
//            folderPrev = folderName;
//          }
//        }

//        if(subFolderName1 !== undefined && subFolderName2 === undefined){
//          if(subFolderName1 !== subFolderPrev1){
//            console.log("SubFolderName1 ",subFolderName1);
//            subFolderPrev1 = subFolderName1;
//          }
//        }

//        if(subFolderName2 !== undefined && subFolderName3 === undefined){
//          if(subFolderName2 !== subFolderPrev2){
//            console.log("SubFolderName2 ",subFolderName1);
//            subFolderPrev2 = subFolderName2;
//          }
//        }

//        if(subFolderName3 !== undefined){
//          if(subFolderName3 !== subFolderPrev3){
//            console.log("SubFolderName3 ",subFolderName3);
//            subFolderPrev3 = subFolderName3;
//          }
//        }

//      } // end of file status check
//    } // end of for loop

    //console.log("folderIDarray="+this.properties.folderNameArray);
    //folderHTML+=`</div></div></div></div>`;
    
    //const listContainer = this.domElement.querySelector('#folderContainer');
    //if(folderContainer){
    //  folderContainer.innerHTML = folderHTML;    
    //}      

    //console.log("F_id=" + this.properties.sf01IDArray);
    //console.log("SF1_id=" + this.properties.sf02IDArray);
    //console.log("SF2_id=" + this.properties.sf03IDArray);

    setTimeout(async ()=> {
      await this.setFolderListeners();
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
*/

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


/*
            // *** Parent Folder ID
            folderID = "dcTab" + tabNum + "-Folder" + fCount;

            if(folderName.replace(/\s+/g, "")===undefined){
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
            this.properties.subFolder1Array.push(subFolderName1);

            if(subFolderName1 !== subFolderPrev1){  
              subFolder01ID = folderID + "-Sub01";

              if(subFolderName1.replace(/\s+/g, "")===undefined){
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
              this.properties.subFolder2Array.push(subFolderName2);                
              
              if(subFolderName2 !== subFolderPrev2){
                console.log("pass x="+x+" sf2="+subFolderName2+" sf2prev="+subFolderPrev2+" sfName2_ID="+subFolderName2ID);
              subFolder02ID = folderID + "-Sub02";

                if(subFolderName2.replace(/\s+/g, "")===undefined){
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
                  if(subFolderName2.replace(/\s+/g, "")===undefined){
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

*/

/*

/*        
        if(folderName !== undefined) {      
          if(folderName !== folderPrev){         

            //if(folderName.replace(/\s+/g, "")===undefined){
            //  folderNameID=folderName;
            //}else{
            //  folderNameID=folderName.replace(/\s+/g, "");
            //}
 
            count=await this.fileCount(results,folderName);
            console.log("count="+count);
            console.log("FolderName ",folderName);

            folderPrev = folderName;            
          }

          if(subFolderName1 !== undefined){
            if(subFolderName1 !== subFolderPrev1){

              //if(subFolderName1.replace(/\s+/g, "")===undefined){
              //  subFolderName1ID=subFolderName1;
              //}else{
              //  subFolderName1ID = subFolderName1.replace(/\s+/g, "");
              //}

              count=await this.fileCount(results,subFolderName1);
              console.log("count="+count);
              console.log("SubFolderName1 ",subFolderName1);

              subFolderPrev1 = subFolderName1;
            }

            if(subFolderName2 !== undefined){
              if(subFolderName2 !== subFolderPrev2){

                //if(subFolderName2.replace(/\s+/g, "")===undefined){
                //  subFolderName2ID=subFolderName2;
                //}else{
                //  subFolderName2ID = subFolderName2.replace(/\s+/g, "");
                //}

                count=await this.fileCount(results,subFolderName2);
                console.log("count="+count);
                console.log("SubFolderName2 ",subFolderName2);
                
                subFolderPrev2 = subFolderName2;
              }

              if(subFolderName3 !== undefined){
                if(subFolderName3 !== subFolderPrev3){

                  //if(subFolderName3.replace(/\s+/g, "")===undefined){
                  //  subFolderName3ID=subFolderName2;
                  //}else{
                  //  subFolderName3ID = subFolderName2.replace(/\s+/g, "");
                  //}

                  count=await this.fileCount(results,subFolderName3);
                  console.log("count="+count);
                  console.log("SubFolderName3 ",subFolderName3);
                  
                  subFolderPrev3 = subFolderName3;                  
                }
              }
            } 
          }
        }
*/        