let request=require("request");
let cheerio=require("cheerio")
let fs=require("fs");
let path=require("path");
let xlsx=require("xlsx");

request("http://www.bpitindia.com/computer-science-and-engineering-faculty-profile.html",whenDataArrive);
function whenDataArrive(err,resp,html){
    let sTool= cheerio.load(html);
    let tableElem=sTool("div.row");
    
    //console.log(tableElem.length);
    let tableElem2=sTool(".col-md-8")
    let tableElem1=sTool(".col-md-4")
    //console.log(tableElem2.leng
    for(let i=0;i<tableElem1.length;i++){
    let FacultyName= sTool(tableElem1[i]).find(".d_inline.fw_600").text().trim();
    console.log(`Name:${FacultyName}`);
        let  Cols=sTool(".table.table-bordered");
            let Cols1=sTool(Cols).find("td");
            let Qualifications= sTool(Cols1[1]).text();
            let Experience= sTool(Cols1[3]).text();
            let Interests= sTool(Cols1[5]).text();
            let Publications= sTool(Cols1[7]).text();
            //console.log(Cols1.length)
           // console.log(`Qualifications: ${Qualifications} Experience: ${Experience} Interests: ${Interests} Publications: ${Publications}`);
           processfaculty(FacultyName,Qualifications,Experience,Interests,Publications);
}}
function processfaculty(FacultyName,Qualifications,Experience,Interests,Publications){
    let Name=FacultyName
    let  profile={
        Name:FacultyName,
        Qualifications:Qualifications,
        Experience:Experience,
        Interests:Interests,
        Publications:Publications,
    }
    if (fs.existsSync(Name)) {
        // file check 
    // console.log("Folder exist")
    }else{
        // create folder 
        // create file
        // add data
        fs.mkdirSync(Name);
    }
    let Name1= path.join(Name,Name+".xlsx");
    let pData=[];
    if(fs.existsSync(Name1)){
     pData=excelReader(Name1,Name)
     pData.push(profile);
    }else{
    // create file
    console.log("File of faculty",Name,"created");
    FData=[profile];
    }
    excelWriter(Name1,FData,Name);
    }
    
    function excelReader(filePath, name) {
        if (!fs.existsSync(filePath)) {
            return null;
        } else{
    // workbook => excel
    let wt = xlsx.readFile(filePath);
    // get data from workbook
    let excelData = wt.Sheets[name];
    // convert excel format to json => array of obj
    let ans = xlsx.utils.sheet_to_json(excelData);
    // console.log(ans);
    return ans;
        }
    }
    function excelWriter(filePath, json, name) {
        // console.log(xlsx.readFile(filePath));
        let newWB = xlsx.utils.book_new();
        // console.log(json);
        let newWS = xlsx.utils.json_to_sheet(json);
        xlsx.utils.book_append_sheet(newWB, newWS, name);  //workbook name as param
    //   file => create , replace
        xlsx.writeFile(newWB, filePath);
    }
    
