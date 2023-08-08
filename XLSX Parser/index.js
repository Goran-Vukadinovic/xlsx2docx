process.env.TZ = "US/Central";

import nodeXlsx from 'node-xlsx';
import * as fs from "fs";
//import { Document,  Packer, Paragraph, TextRun} from "docx";
import builder from 'docx-builder';
import os from 'os';


var dir = './data';
if (!fs.existsSync(dir)){
    fs.mkdirSync(dir, { recursive: true });
}

dir = './data/2021';
if (!fs.existsSync(dir)){
    fs.mkdirSync(dir, { recursive: true });
}

dir = './data/2022';
if (!fs.existsSync(dir)){
    fs.mkdirSync(dir, { recursive: true });
}

// Parse a file
const wb = nodeXlsx.parse(`./2021-2022 inquiries.xlsx`, {cellDates: true, dateNF:"dd/mm/yy"});
var ws = wb[0];
//console.log(JSON.stringify(ws));
var sheetName = ws.name;
var data = ws.data;

var aryTitle = {};
for(var i = 1 ; i < data.length ; i++)
{
    //if(i != 579) continue;
    //if(i != 428 && i != 309) continue;
    var line = data[i];
    if(line.length != 6)
    {
        line.push("");
    }
    if(line.length > 6)
    {
        console.log("000000000000000000");
        break;
    }
    var title = escapeHtml(line[0]);
    var recvDate = line[1];
    var question = escapeHtml(line[2]);
    var response = escapeHtml(line[3]);
    var category = escapeHtml(line[4]);
    var topicWord = escapeHtml(line[5]);

    var t = title.replaceAll("/", "_");
    t = t.replaceAll("\\", "_");
    t = t.replaceAll("?", "");    
    t = t.replaceAll("*", "");
    t = t.replaceAll("\"", "'");
    t = t.replaceAll("<", "(");
    t = t.replaceAll(">", ")");
    t = t.replaceAll("|", "_");
    t = t.replaceAll(":", "_");
    t = t.trim();
    var tKey = t.toUpperCase();
    var prev = aryTitle[tKey];
    if(!prev){
        aryTitle[tKey] = 1
    } 
    else{                
        t += "_" + aryTitle[tKey];
        aryTitle[tKey] = aryTitle[tKey] + 1;        
    }
    
    var docx = new builder.Document();
    docx.unsetBold();
    docx.unsetItalic();
    docx.unsetUnderline();
    docx.setFont("Verdana");
    docx.setSize(20);
    docx.leftAlign();
    docx.insertText(title);
    docx.insertText(getDateString(recvDate));
    docx.setBold();
    docx.insertText("Question 1:");
    docx.unsetBold();    
    if(question){
        var qd = question.split("\n");
        for(var j = 0 ; j < qd.length ; j++)
        {
            if(qd[j]) docx.insertText(qd[j]);
        }
    }
    else{
        console.log(i);
        docx.insertText("");
    }
    
    docx.setBold();
    docx.insertText("Answer 1:");
    docx.unsetBold();
    if(response)
    {
        var qr = response.split("\n");
        for(var j = 0 ; j < qr.length ; j++)
        {
            if(qr[j]) docx.insertText(qr[j]);
        }
    }
    else{
        docx.insertText("");
        console.log(i);
    }
    var path = "";
    if(recvDate.getFullYear() == 2021)
    {
        path = "./data/2021/" + t + ".docx";
    }
    else if(recvDate.getFullYear() == 2022)
    {
        path = "./data/2022/" + t + ".docx";
    }
    else{
        path = "./data/" + t + ".docx";
    }


    docx.save(path, function(err){
        if(err) console.log(err);
    });
}

function getDateString(d)
{
    const monthNames = ["January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
        ];
        return monthNames[d.getMonth()] + " " + d.getDate() + ", " + (d.getFullYear());
}
/*
new TextRun(recvDate + ""),
                            new TextRun({
                                text: "Question 1:",
                                bold: true,
                            }),
                            new TextRun(question),
                            new TextRun({
                                text: "Answer 1:",
                                bold: true,
                            }),
                            new TextRun(response),
*/
function saveDocx(doc, fName)
{
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("./data/" + fName + ".docx", buffer);
    });
}

/*
docx.setBold();
    docx.unsetBold();
    docx.setItalic();
    docx.unsetItalic();
    docx.setUnderline();
    docx.unsetUnderline();
    docx.setFont("Verdana");
    docx.setSize(40);
    docx.rightAlign();
    docx.centerAlign();
    docx.leftAlign();
    docx.insertText("Hello");
*/

function escapeHtml(unsafe)
{
    if(!unsafe) return "";
    return unsafe
         .replace(/&/g, "&amp;")
         .replace(/</g, "&lt;")
         .replace(/>/g, "&gt;")
         .replace(/"/g, "&quot;")
         .replace(/'/g, "&#039;");
 }