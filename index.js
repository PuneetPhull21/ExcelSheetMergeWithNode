const reader = require('xlsx');
const express = require('express');
const multer = require('multer');
const bodyparser =require('body-parser');
var app = express();
app.use(bodyparser.json());
app.use(bodyparser.urlencoded({ extended: true }));

var storage = multer.diskStorage({
    destination:(req,file,cb)=>{
        cb(null,'./UploadedSheets');
    },
    filename:(req,file,cb)=>{
        
        cb(null,file.fieldname+""+Date.now()+""+file.originalname);
        
    }
})

var upload = multer({storage:storage}).array('Sheets',2);


app.post('/upload',function(req,res){
    upload(req,res,function(err){
        if(err){
         console.log(err)   
        }
const file = reader.readFile(`./UploadedSheets/${req.files[0].filename}`);
const file2 = reader.readFile(`./UploadedSheets/${req.files[1].filename}`);
const file3 = reader.readFile('./sheet3.xls');

let data = [];
let data2 = [];
let finaldata = [];
const sheets = file.SheetNames;
for(let i=0;i<sheets.length;i++){
    const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]]);
 const temp2 = reader.utils.sheet_to_json(file2.Sheets[file.SheetNames[i]]);
 temp.forEach((res)=>{
     data.push(res);
 })
 temp2.forEach((res)=>{
     data2.push(res);
 })
}
for(let j=0;j<data.length;j++){
    for(let k=0;k<data2.length;k++){
        if(data[j].Email === data2[k].Email){
            const final = {
                EmployeeId:data[j].EmployeeId,
                EmployeeNAme:data[j].EmployeeNAme,
                Email:data[j].Email,
                Date:data[j].Date,
                Hours:data[j].Hours,
                LeaveStatus:data2[k].LeaveStatus,
            }
            finaldata[j]=final;
        }
        
    }
}

const ws = reader.utils.json_to_sheet(finaldata);
  
reader.utils.book_append_sheet(file3,ws,"Sheet7");

reader.writeFile(file3,'./sheet3.xls');

console.log("Download in the sheet ");
         
        res.json({
            status:200,
            success:true,
            data:finaldata
        });
    })
})

app.listen(3000,function(){
    console.log("server is running on port 3000");
})

