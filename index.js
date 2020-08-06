/*
    WIE Telegram Bot Alpha Build V 0.5
    Programmer: Lance CHAN Wan Lung
    Advisor: Adam Wong

    Special thanks to Yagop for building node-telegram-bot-api
    Framework: node.js
    Language: Javascript
    Database: MongoDB
    Data Visualization: HighCharts 
*/

/*
============================================TO DO LIST=============================================

[*] 1. Fix the User State Iteration.
[] 2. Simplify statement objects and variables.
[] 3. Await for survey documentation and put it into json and read from it as file (.json).

 */



// API and database runtime preparation

const TelegramBot = require('node-telegram-bot-api');
const MongoClient = require('mongodb').MongoClient;
const fs = require('fs');
var XLSX = require('xlsx');
const chartExporter = require("highcharts-export-server");

// JSON config pre-load

const Config = require('./json/tg_config.json');
const Sticker = require('./json/sticker-lookup.json');
const ReplyKeyboard = require('./json/reply_markup.json');
const Survey = require('./json/survey.json');

// Object constant for following code

const bot = new TelegramBot(Config.token, {polling: true});
const client = new MongoClient(Config.uri, { useUnifiedTopology: true });

//User variables

const error = Config.errormsg;
var mapSP = Config.mapPicture;
var buttonmsg = Config.buttonmsg;
var surveyTimer = 5000; //1000 = 1 second
var updateTimer = 5000;


client.connect(err =>{
    
    bot.onText(/\/start/, (msg) => {
        bot.sendMessage(msg.chat.id,"Hi, "+msg.from.first_name);
        var msgdate = new Date(msg.date*1000);
        var lastvisit = "";
        var userobj = {tgname: msg.from.first_name, date: Date(msgdate)};
        const collection = client.db("WIE_TGUser").collection("WIE_SPEEDTG");
        var findname = {tgname: msg.from.first_name};
        var updateobj = {$set:userobj};
        collection.find(findname).toArray(function(err, result) {
            if (err) throw err;
            var search = result[0];
            console.log(search);
            if(!search){
                collection.insertOne(userobj);
                console.log("1 document inserted");
                bot.getMe().then((bots)=>{
                    var intro = Config.intromsg.intro1 + bots.first_name + Config.intromsg.intro2;
                    bot.sendMessage(msg.chat.id,intro + "\n\nYour data has been registered to the Chatbot!\n\nWe will mark your user name and your last visited time here, Enjoy!",ReplyKeyboard);
                })
            } else {
                var lastdate = new Date(result[0].date);
                var timediff = Math.abs(msgdate.getTime() - lastdate.getTime()+2000);
                lastvisit += "It is nice to see you back.  You were last here at "+result[0].date+".\n\n"+tellTimeDiffLastTime(timediff)+"\n\nTo refresh your memory, here is my introduction again.\n\n";
                collection.updateOne(findname, updateobj);
                console.log("1 document updated");
                console.log("Updated Date:" +Date(msgdate) +"(UNIX : "+(msg.date*1000)+")");
                bot.getMe().then((bots)=>{
                    var intro = Config.intromsg.intro1 + bots.first_name + Config.intromsg.intro2;
                    bot.sendMessage(msg.chat.id,lastvisit + intro,ReplyKeyboard);
                })
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
                "Current Time 現在時間: " + time +desmsg+ "\n\n" + t +limitmsg + mnmsg+"\n\n**Actual Time of Arrival subjects to the traffic on the road and the arrangement from the facilites.\n**實際到達時間需要視路面情況及院校安排而定。");
            bp1 = false;
            bp2 = false;
        });
    })
    bot.on('message', (msg) => {
        var initcheck = Boolean(/\/start/.test(msg.text) || /shuttle/i.test(msg.text));
        if(!initcheck){
            var strcheck = checkMessageAndSendback(msg.chat,msg.text);
            if(!strcheck){
                bot.sendMessage(msg.chat.id,error);
                bot.sendSticker(msg.chat.id,Sticker.No);
            } else {
                checkSurveyVaild(msg);
            }
        }
    });

//========================================
// Survey and score point function series
//========================================

    function checkSurveyVaild(msg){
        inlineSurveyReminder(msg);
        
        
        // !!!!!!!!!!!!!! WILL USE LATER DON'T DELETE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!


        /* const collection = client.db("WIE_TGUser").collection("WIE_SPEEDTG");
        var cursor = {tgname: msg.from.first_name};
        collection.find(cursor).toArray(function (err,result){
            if (err) throw err;
            if(!result[0].Q1){
                console.log("User did not fill in the survey, Initiaite the survey.");inlineSurveyReminder(msg);
            }
        }); */
    }
    var st;
    var scorearr = [];
    function inlineSurveyReminder(msg){
        if(st) clearTimeout(st);
        st = setTimeout(function(){
            bot.sendMessage(msg.chat.id,"Hey, would you like to help me to fill out some survey?",Survey.ikyn);
            bot.on('callback_query', query=>{
                var scorestate = {
                    chat:query.message.chat.id,
                    username:query.from.first_name,
                    q1:null,
                    q2:null,
                    q3:null
                }
                if(!scorearr[0]){
                    scorearr[0] = scorestate;
                    console.log("Creating Survey Buffer.");
                } else {
                    checkLoop:
                        for(var i = 0;i<scorearr.length;i++){
                            if(scorearr[i].chat == query.message.chat.id){
                                break checkLoop;
                            } else if(scorearr[i].chat != query.message.chat.id && i == (scorearr.length-1)){
                                scorearr[scorearr.length] = scorestate;
                                console.log("Create User "+(scorearr.length-1)+".");
                                break checkLoop;
                            }
                        }
                    }
                if(query.data == "survey_reject"){
                    bot.editMessageText("That's Okay, see you next time~",{chat_id:query.message.chat.id,message_id:query.message.message_id});
                    query.data = null;
                }
                if(query.data == "survey_accept"){
                    bot.editMessageText(Survey.question[0],{chat_id:query.message.chat.id,message_id:query.message.message_id,reply_markup:Survey.ikscore.reply_markup});
                    query.data = null;
                }
                for(var i = 0;i<scorearr.length;i++){
                    if(query.message.chat.id == scorearr[i].chat){
                        console.log('User '+i+' Entering Survey.');
                        if(query.data && query.message.text == Survey.question[0]){
                            bot.editMessageText(Survey.question[1],{chat_id:query.message.chat.id,message_id:query.message.message_id,reply_markup:Survey.ikscore.reply_markup});
                            scorearr[i].q1 = parseInt(query.data);
                            query.data = null;
                        }
                        if(query.data && query.message.text == Survey.question[1]){
                            bot.editMessageText(Survey.question[2],{chat_id:query.message.chat.id,message_id:query.message.message_id,reply_markup:Survey.ikscore.reply_markup});
                            scorearr[i].q2 = parseInt(query.data);
                            query.data = null;
                        }
                        if(query.data && query.message.text == Survey.question[2]){
                            bot.editMessageText("Done",{chat_id:query.message.chat.id,message_id:query.message.message_id});
                            scorearr[i].q3 = parseInt(query.data);
                            var scoresheet = [scorearr[i].q1,scorearr[i].q2,scorearr[i].q3];
                            console.log(scoresheet);
                            scoresheetToDB(query.from.first_name,scoresheet);
                            query.data = null;
                        }
                    }
                } 
            });
        },surveyTimer);
    }
    function scoresheetToDB(tgname,score){
        const collection = client.db("WIE_TGUser").collection("WIE_SPEEDTG");
        var cursor = {tgname: tgname};
        var updateobj = {$set:{Q1:score[0],Q2:score[1],Q3:score[2],Q4:score[3]}}
        collection.find(cursor).toArray(function (err,result){
            if (err) throw err;
            collection.updateOne(cursor, updateobj);
            console.log("Updated User Survey Result.");
        });
    }
    
//========================================
//      Update DB to Text and Excel
//========================================

    function updateUserDBToText(){
        const collection = client.db("WIE_TGUser").collection("WIE_SPEEDTG");
        var file = createExcelFile();
        var userwb = file[0];
        var userws = file[1];
        if (fs.existsSync('WIE_Records.xlsx')) {
            userwb = XLSX.readFile('WIE_Records.xlsx');
            userws = userwb.Sheets["User Record"];
            userws2 = userwb.Sheets["Choice Chart"];
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
            userws['!cols'] = [{wpx:220},{wpx:330},{wpx:360}];
            userwb.Sheets["User Record"] = userws;
            XLSX.writeFileAsync('WIE_Records.xlsx',userwb,{bookType:"xlsx"},(err) =>{
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
        userwb.SheetNames.push("Choice Chart");
        var cl_data = [["Name","Last Visited Date","Recent Choice"]];
        var userws = XLSX.utils.aoa_to_sheet(cl_data);
        userws['!cols'] = [{wpx:220},{wpx:330},{wpx:360}];
        return [userwb,userws];
    }
    function exportDataToChartImage(){
        const collection = client.db("WIE_TGUser").collection("WIE_LOG");
        var ydata = new Array(0,0,0,0,0,0);
        var xdata = buttonmsg;
        collection.find({},{projection:{_id:0,choice:1}}).toArray(function (err,result){
            if (err) throw err;
            for(var y = 0; y < buttonmsg.length;y++){
                for(var i = 0; i < result.length;i++){
                    if(result[i].choice == buttonmsg[y]){
                        ydata[y] += 1;
                    }
                }
            }
            makeChartOnResult(xdata,ydata);
        });
    }
    function makeChartOnResult(label,value){
        // Initialize the exporter
        chartExporter.initPool();
        // Chart details object specifies chart type and data to plot
        var data = [];
        var chartDetails = require('./json/chart_format.json').chartDetail;
        for(var i=0;i<value.length;i++){
            data.push({name:label[i],y:value[i]});
        }
        if(data[0]){
            chartDetails.options.series[0] = {data:data};
        } else if(!data[0]){
            return false;
        }
        chartExporter.export(chartDetails, (err, res) => {
            if(err) return console.log(err);
            // Get the image data (base64)
            let imageb64 = res.data;
            // Filename of the output
            let outputFile = "bar.png";
            // Save the image to file
            fs.writeFileSync(outputFile, imageb64, "base64", function(err) {
                if (err) console.log(err);
            });
                chartExporter.killPool();
        });
    }

//========================================
//      Chatbot Interaction Markdown
//========================================

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
    function checkMessageAndSendback(chat,msgstr){
        var kbtxt = ReplyKeyboard.reply_markup.keyboard;
        console.log(kbtxt[3]);
        if (msgstr == kbtxt[3][0]){
            fs.readFile('SPEED_place.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMediaGroup(chat.id,mapSP);
                bot.sendMessage(chat.id,data);
            })
            return true;
        }
        if (msgstr == kbtxt[0][0] || /what jobs/ig.test(msgstr)){
            fs.readFile('./text/text_a.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                msgstr = buttonmsg[0] ;
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == kbtxt[0][1] || /Looking for jobs/ig.test(msgstr)){
            fs.readFile('./text/text_b.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.addoil);
                msgstr = buttonmsg[1];
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == kbtxt[1][0] || /Got offer/ig.test(msgstr)){
            fs.readFile('./text/text_c.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.XD);
                msgstr = buttonmsg[2] ;
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == kbtxt[1][1] || /Employed/ig.test(msgstr)){
            fs.readFile('./text/text_d.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.food);
                msgstr = buttonmsg[3];
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == kbtxt[2][0] || /Got 300 hrs/ig.test(msgstr)){
            fs.readFile('./text/text_e.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                bot.sendSticker(chat.id,Sticker.heart);
                msgstr = buttonmsg[4];
                markChoice(chat,msgstr);
            })
            return true;
        }
        if (msgstr == kbtxt[2][1] || /Where are the forms/ig.test(msgstr)){
            fs.readFile('./text/text_f.txt', 'utf-8', (err, data) => { 
                if (err) throw err; 
                bot.sendMessage(chat.id,data);
                msgstr = buttonmsg[5];
                markChoice(chat,msgstr);
            })
            return true;
        }
    }
    function markChoice(chat,msgstr){
        const collection = client.db("WIE_TGUser").collection("WIE_LOG");
        var findname = {tgname:chat.first_name};
        var userobj = {tgname:chat.first_name,choice:msgstr};
        var updateobj = {$set:userobj};
        collection.find(findname).toArray(function(err, result) {
            if (err) throw err;
            var search = result[0];
            console.log(search);
            if(!search){
                collection.insertOne(userobj);
                console.log("WIE_LOG: 1 document inserted");
            } else {
                collection.updateOne(findname,updateobj);
                console.log("WIE_LOG: Updated "+chat.first_name+"'s choice and printed on the file.");
            }
        });
    }
    bot.on('polling_error',(err)=>{
        console.log(err);
    })

// Interval List
    setInterval(updateUserDBToText,updateTimer);
    setInterval(exportDataToChartImage,10000);
});