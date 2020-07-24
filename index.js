/*
    WIE Telegram Bot Alpha Build V 0.5
    Programmer: Lance CHAN Wan Lung
    Advisor: Adam Wong

    Special thanks to Yagop for building node-telegram-bot-api
    Framework: node.js
    Language: Javascript
    Database: MongoDB
*/

/*
[*]1. Fix login date on last visited.
[*]2. Re adjust the msg for intro.
[]3. Remember the choices and put it into spreadsheet report(Better if can be recorded to Excel file).
[]4. Survey on inline message with integer and write on database with name and result (will later documented) ,while the user actually pressed the reply markup recently.

 */



// API and database runtime preparation
const TelegramBot = require('node-telegram-bot-api');
const MongoClient = require('mongodb').MongoClient;
var XLSX = require('xlsx');

// JSON config pre-load
const Config = require('./tg_config.json');
const Sticker = require('./sticker-lookup.json');
const ReplyKeyboard = require('./reply_markup.json');
const token = Config.token;
const uri = Config.uri;

const fs = require('fs');

// Object constant for following code
const bot = new TelegramBot(token, {polling: true});
const client = new MongoClient(uri, { useUnifiedTopology: true });


// *** String on start, Can be edited. ***
var botname = "WIE_SPEED Orientation Chatbot";
var intro = "My name is "+botname+".  I am created to help you with WIE. I will try to answer your questions regarding WIE. Click one of the buttons to ask me a question. At any time, you can show the buttons again by clicking the keyboard icon next to the smiley on the bottom right hand corner.\n你好，歡迎你使用PolyU SPEED WIE Telegram智能對話機械人!\n你可以使用輸入框以下的按鍵去選擇你想向機械人查詢的資訊！";

//Unused
const error = "Sorry, but I have no idea what you are talking about.\n\n 對不起，我不能理解你這句的意思。";

//User variables

var mapSP = [
    {
        "type":"photo", 
        "media":"https://www.speed-polyu.edu.hk/f/page/1316/1897/20161213_PolyU-WK_map.jpg"
    },
    {
        "type":"photo", 
        "media":"https://www.speed-polyu.edu.hk/f/page/1316/7733/HHB%20Map%20(New).jpg"
    },
    {
        "type":"photo", 
        "media":"https://www.speed-polyu.edu.hk/f/page/1316/1897/(20140130)%20PolyU%20Main%20Campus.png"
    }
];

client.connect(err =>{
    function tellTimeDiffLastTime(diff){
        var td = {
            second : Math.floor(diff / 1000),
            minute : Math.floor(diff / (1000*60)),
            hour : Math.floor(diff  / (1000*60*60)),
            day : Math.floor(diff / (1000*60*60*24))
        };
        console.log(td);
        var diffmsg = "You have been leaved the chatbot for "+td.day;
        if(td.day == 0 || td.day > 1){
            diffmsg += " days ";
        } else if (td.day == 1){
            diffmsg += " day ";
        }
        if(td.hour%24 == 0 || td.hour%24 > 1){
            diffmsg += td.hour%24 +" hours ";
        } else if(td.hour%24 == 1){
            diffmsg += td.hour%24 +" hour ";
        }

        if(td.minute%60 == 0 || td.minute%60 > 1){
            diffmsg += td.minute%60 +" minutes ";
        } else if(td.minute%60 == 1){
            diffmsg += td.minute%60 +" minute ";
        } 

        if(td.second%60 == 0 || td.second%60 > 1){
            diffmsg += td.second%60 +" seconds.";
        } else if(td.second%60 == 1){
            diffmsg += td.second%60 +" second.";
        }
        return diffmsg;
    }
    bot.onText(/\/start/, (msg) => {
        bot.sendMessage(msg.chat.id,"Hi, "+msg.from.first_name);
        var msgdate = new Date(msg.date*1000);
        var lastvisit = "";
        var userobj = {
            tgname: msg.from.first_name, 
            date: Date(msgdate),
            choice: ""
        };
        const collection = client.db("WIE_TGUser").collection("WIE_SPEEDTG");
        var findname = {tgname: msg.from.first_name};
        var updateobj = {$set:{
            tgname: msg.from.first_name, 
            date: Date(msgdate),
        }};
        collection.find(findname).toArray(function(err, result) {
            if (err) throw err;
            var search = result[0];
            console.log(search);
            if(!search){
                collection.insertOne(userobj);
                console.log("1 document inserted");
                bot.sendMessage(msg.chat.id,intro + "\n\nYour data has been registered to the Chatbot!\n\nWe will mark your user name and your last visited time here, Enjoy!",ReplyKeyboard);
            } else {
                var lastdate = new Date(result[0].date);
                var timediff = Math.abs(msgdate.getTime() - lastdate.getTime() -1000);
                lastvisit += "It is nice to see you back.  You were last here at "+result[0].date+".\n\n"+tellTimeDiffLastTime(timediff)+"\n\nTo refresh your memory, here is my introduction again.\n\n";
                collection.updateOne(findname, updateobj);
                console.log("1 document updated");
                console.log("Updated Date:" +Date(msgdate) +"(UNIX : "+(msg.date*1000)+")");
                bot.sendMessage(msg.chat.id,lastvisit + intro,ReplyKeyboard);
            }
        });
        bot.sendSticker(msg.chat.id, Sticker.write);
    });
    bot.onText(/shuttle/i, (msg) =>{
        var today = new Date(msg.date*1000);
        var place = /WK|HHB/i;
        var WK = /WK/ig;
        var HHB = /HHB/ig;
        var h = today.getHours();
        var m = today.getMinutes();
        if(m<10){
            m = "0"+m.toString();
        }
        if(h<10){
            h = "0"+h.toString();
        }
        var time = h+":"+m;
        
        var cursor = {venue:"HHB_WK"};
        var des = "West Kowloon (Yau Ma Tei) Campus 西九龍(油麻地)校園";
        var bp1 = false;
        var bp2 = false;
        var iHHB,iWK;
        while(WK.test(msg.text) == true){
            iWK = WK.lastIndex;
        }
        while(HHB.test(msg.text) == true){
            iHHB = HHB.lastIndex;
        }
        if(place.test(msg.text)){
            console.log(time);
            if(( (iHHB<iWK) || (iHHB && !iWK) && bp1 == false)){
                cursor = {venue:"WK_HHB"};
                des = "Hung Hom Bay Campus 紅磡灣校園";
                bp1 = true; 
                if(iWK && !bp2){
                    cursor = {venue:"HHB_WK"};
                    des = "West Kowloon (Yau Ma Tei) Campus 西九龍(油麻地)校園";
                    bp2 = true;
                }
            }
            if(((iHHB>iWK) || (!iHHB && iWK) && bp1 == false)){
                cursor = {venue:"HHB_WK"};
                des = "West Kowloon (Yau Ma Tei) Campus 西九龍(油麻地)校園";
                bp1 = true;
                if(iHHB && !bp2){
                    cursor = {venue:"WK_HHB"};
                    des = "Hung Hom Bay Campus 紅磡灣校園";    
                    bp2 = true;
                }
            }
        }
        const shuffle = client.db("WIE_TGUser").collection("WIE_Shuffle");
        shuffle.find(cursor).toArray(function (err, result) {
            if (err) throw err;
            var tobj = result[0].time;
            var i,t = "Shuttle Bus Timetable 穿梭巴士時間表\n\n" , lim = 0;
            var desmsg = "\nDestination 目的地：\n" +des; 
            for(i = 0;i<tobj.length;i++){
                if(lim <3){
                    if(time<=tobj[i].toString()){
                        var rm,rh;
                        rm = parseInt(tobj[i].substr(3,2) - time.substr(3,2));
                        rh = parseInt(tobj[i].substr(0,2) - time.substr(0,2));
                        if(rm < 0){
                            rm = 60 + rm; // rm is negative so need to plus
                            rh = rh - 1;
                        }
                        t += tobj[i] + "\n";
                        t += "Remaining Time 剩餘時間: " +rh+" Hours "+rm+ " Minutes.\n\n";
                        lim++;
                    }
                }
            }
            var limitmsg = "\nOnly show " + lim + " upcoming bus.\n只顯示 "+ lim +" 班巴士時間。";
            var mnmsg = "\nService Period: Monday to Friday, except Public Holidays.\n服務時間：星期一至五，公眾假期除外。"
            if(lim == 0){
                limitmsg = "\nCurrently there is no bus service.\n現時並沒有巴士服務。";
            }
            
            bot.sendMessage(msg.chat.id, 
                "Current Time 現在時間: " + time +desmsg+ "\n\n" + t +limitmsg + "\n\n**Actual Time of Arrival subjects to the traffic on the road and the arrangement from the facilites.\n**實際到達時間需要視路面情況及院校安排而定。");
            bp1 = false;
            bp2 = false;
        });
    })
    function checkMessageAndSendback(chat,msgstr){
        if (msgstr == "Where can I find PolyU SPEED?\nPolyU SPEED 在哪裡？"){
            fs.readFile('SPEED_place.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMediaGroup(chat.id,mapSP);
                bot.sendMessage(chat.id,data);
            })
            return true;
        }
        if (msgstr == "A. What jobs are suitable for WIE?" || /what jobs/ig.test(msgstr)){
            fs.readFile('./text/text_a.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                msgstr = "A. What jobs are suitable for WIE?" ;
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == "B. I have no or less than 300 hrs job experience. \nI am looking for a job." || /Looking for jobs/ig.test(msgstr)){
            fs.readFile('./text/text_b.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.addoil);
                msgstr = "B. I have no or less than 300 hrs job experience. I am looking for a job.";
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == "C. I have no or less than 300 hrs job experience. \nI just received an offer." || /Got offer/ig.test(msgstr)){
            fs.readFile('./text/text_c.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.XD);
                msgstr = "C. I have no or less than 300 hrs job experience. I just received an offer." ;
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == "D. I have no or less than 300 hrs job experience. \nI am already employed." || /Employed/ig.test(msgstr)){
            fs.readFile('./text/text_d.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.food);
                msgstr = "D. I have no or less than 300 hrs job experience. I am already employed." ;
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == "E. I have 300 hrs, or more, of job experience." || /Got 300 hrs/ig.test(msgstr)){
            fs.readFile('./text/text_e.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.heart);
                msgstr = "E. I have 300 hrs, or more, of job experience." ;
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == "F. Where can I find the forms?" || /Where are the forms/ig.test(msgstr)){
            fs.readFile('./text/text_f.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                msgstr = "F. Where can I find the forms?" ;
                markChoice(chat,msgstr);
            })
            return true;
        }
    }
    function markChoice(chat,msgstr){
        const collection = client.db("WIE_TGUser").collection("WIE_SPEEDTG");
        var findname = {tgname: chat.first_name};
        var updateobj = {$set:{choice: msgstr}};
        collection.find(findname).toArray(function(err, result) {
            if (err) throw err;
            var search = result[0];
            console.log(search);
            if(!search){
                bot.sendMessage(chat,"Error Occured, please re-launch the chatbot.");
            } else {
                collection.updateOne(findname,updateobj);
                console.log("Updated "+chat.first_name+"'s choice and printed on the file.");
            }
        });
    }
    bot.on('message', (msg) => {
        var initcheck = Boolean(/\/start/.test(msg.text.toString()) || /shuttle/i.test(msg.text.toString()));
        if(!initcheck){
            var strcheck = checkMessageAndSendback(msg.chat,msg.text.toString());
            if(!strcheck){
                bot.sendMessage(msg.chat.id,error);
                bot.sendSticker(msg.chat.id,Sticker.No);
            }
        }
    });
    setInterval(updateUserDBToText,5000);

    function updateUserDBToText(){
        const collection = client.db("WIE_TGUser").collection("WIE_SPEEDTG");
        var file = createExcelFile();
        var userwb = file[0];
        var userws = file[1];
        if (fs.existsSync('WIE_Records.xlsx')) {
            userwb = XLSX.readFile('WIE_Records.xlsx');
            userws = userwb.Sheets["User Record"];
        }
        collection.find({}).toArray(function(err,result){
            if (err) throw err;
            fs.writeFile('usertext.txt',JSON.stringify(result),function(err){
                if (err) throw err;
            });
            var data = [];
            for(var i=0;i<result.length;i++){
                data.push({
                    "Name": result[i].tgname,
                    "Last Visited Date": result[i].date,
                    "Recent Choice": result[i].choice
                });
            }
            userws = XLSX.utils.json_to_sheet(data,{header:["Name","Last Visited Date","Recent Choice"]});
            userws['!cols'] = [{wch:360}];
            userwb.Sheets["User Record"] = userws;
            XLSX.writeFileAsync('WIE_Records.xlsx',userwb,(err) =>{
            });
        })
    }

    function createExcelFile(){
        var userwb = XLSX.utils.book_new();
        userwb.Props = {
            Title: "WIE User Database",
            Subject: "Database",
            Author: "SPEED WIE Chatbot",
            CreatedDate: new Date()
        }   
        userwb.SheetNames.push("User Record");
        var cl_data = [["Name","Last Visited Date","Recent Choice"]];
        var userws = XLSX.utils.aoa_to_sheet(cl_data);
        return [userwb,userws];
    }

});