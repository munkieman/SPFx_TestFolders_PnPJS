export interface IDocumentCentreProps {
    description: string;
    
    URL:string;
    tenantURL: any;
    dcURL: string;   
    siteName: string;
    siteShortName: string;
    siteTitle: string;
    divisionName: string;
    divisionTitle: string;
    teamName: string;
    libraryName: string;
    
    divisions: ["Assessments","Central","Connect","Employability","Health"];
    dataResults: any[];
    siteArray: any;
    folderArray: string[];
    subFolder1Array: string[];
    subFolder2Array: string[];
    subFolder3Array: string[];
    sf01IDArray: string[];
    sf02IDArray: string[];
    sf03IDArray: string[];
    filesArray : any[];
    dataFlag : boolean;

    getFilesCallFlag:boolean;
    isDCPowerUser:boolean;
    asmPowerUser:boolean;
    cenPowerUser:boolean;
    cnnPowerUser:boolean;
    empPowerUser:boolean;
    heaPowerUser:boolean;
    isManager:boolean;
    powerUserToolsHTML:string;
  
    asmGroupID:string;
    cenGroupID:string;
    cnnGroupID:string;
    empGroupID:string;
    heaGroupID:string;
    managersGroupID:string;
  
    name: string;
    Slider:string;
    Toggle:string;
    dropdowm:string;
    checkbox:string;
    textbox:string;
  }
  