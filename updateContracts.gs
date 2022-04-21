var ssID = "XXXX";

function onOpen(e){

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Панель пользователя')
      .addItem('Ввод нового клиента', 'newClient')
      .addItem('Ввод нового города', 'newTowns')  
      .addItem('Ввод нового вида работ', 'newWorks')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Загрузка данных')
         .addItem('Прогрузить Клиентов', 'loadClients')
         .addItem('Прогрузить Города', 'loadTowns')
         .addItem('Прогрузить Вид работ','loadWorks'))
      .addToUi();
}

function Access(){
  url = "https://script.google.com/macros/s/AKfycbxWxZzKIuHqtmc1Gm9l5RvAMimOEXO4eiS20Vu_JsYPlw8pBEM/exec";
  var res = UrlFetchApp.fetch(url);
}

function noAccess(){
  url = "https://script.google.com/macros/s/AKfycbz8luwCqsAFXalZzQdLi_YYxitYdgwLX96OxT5G9W2Zl77Nf5Q/exec";
  var res = UrlFetchApp.fetch(url);
}

function SortWorks(){
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Вид работ");
  var editedCell = wsData.getActiveCell();
  
  var columnToSortBy = 1;

  if(editedCell.getColumn() == columnToSortBy){   
    var range = wsData.getRange(5, 1, wsData.getLastRow(), wsData.getLastColumn());
    range.sort( { column : columnToSortBy, ascending: true } );
  }
}
function SortTowns(){
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Города");
  var editedCell = wsData.getActiveCell();
  
  var columnToSortBy = 1;

  if(editedCell.getColumn() == columnToSortBy){   
    var range = wsData.getRange(2, 1, wsData.getLastRow(), wsData.getLastColumn());
    range.sort( { column : columnToSortBy, ascending: true } );
  }
}

function noProtect(funcId, data, data2, list, chk){
  var parameters = "?ssID="+ssID;
  var url = "https://script.google.com/macros/s/AKfycbzwltMuJOUypforoS3oc_RcFZwUiBKV5xrtx-M9NzenKEEayCnP/exec"+parameters;
  var res = UrlFetchApp.fetch(url);
  var response = res.getResponseCode();
  if (response == 200 )return[5, true, data, data2, list, chk]; 
  else return [5, false, data, data2, list, chk];
}

function Protect(){
  var parameters = "?ssID="+ssID;
  var url = "https://script.google.com/macros/s/AKfycbzzQZLTXJypFHBtnQ7ImQi-LBUNI69Zwzj750nhZFrJuFrl97c/exec"+parameters;
  var res = UrlFetchApp.fetch(url);
  var response = res.getResponseCode();
  if (response == 200 )return[6, true]; 
  else return [6, false];
}

function CheckAccess(){
  let user = PropertiesService.getUserProperties().getProperties();
  user = Session.getActiveUser().getEmail().toLowerCase();
  let sheet = SpreadsheetApp.openById(ssID).getSheetByName('Тех. данные');
  let accessList = sheet.getRange(2, 11, 30).getValues().map(item => {return item[0];}).filter(item => {return item !== "" ;});
  let access = false;
  accessList.forEach(item => {
    if (user == item){
      access = true;
    }
  });
  if (access == false){
    SpreadsheetApp.getUi().alert('Отказано в доступе! Запросите разрешение у владельца таблицы');
    return false;
  }
  return true;
}

function Access(){
  var access = getUsers();
  if (access[1] === undefined){
    SpreadsheetApp.getUi().alert('Отказано в доступе! Запросите разрешение у владельца таблицы');
    return false;
  }
  return true;
}

function newClient(){
  if (!Access()) return false;
  var html = HtmlService.createHtmlOutputFromFile('modalNewClient.html');
  var ui = SpreadsheetApp.getUi();
  html.setWidth(610);
  html.setHeight(900);
  ui.showModalDialog(html, 'Форма ввода нового клиента');
}

function newTowns(){
  if (!Access()) return false;
  var html = HtmlService.createHtmlOutputFromFile('modalNewTowns.html');
  var ui = SpreadsheetApp.getUi();
  html.setWidth(610);
  html.setHeight(500);
  ui.showModalDialog(html, 'Форма ввода нового города');
}

function newWorks(){
  if (!CheckAccess()) return false;
  var html = HtmlService.createHtmlOutputFromFile('modalNewWorks.html');
  var ui = SpreadsheetApp.getUi();
  html.setWidth(610);
  html.setHeight(500);
  ui.showModalDialog(html, 'Форма ввода нового вида работ');
}

function getData(funcId, ssName, id, column, idlabel){
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssName);

  var data = wsData.getRange(2,column, wsData.getLastRow() - 1).getValues().map(function(o){ return o[0];}).filter(function(o){ return o !== "";});
  var label = wsData.getRange(1,column).getValues()[0];
  label = label[0]; 
  return [funcId, data, id, label, idlabel];
}

function getDataCompany(funcId, ssName, id, column, idlabel, category){
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssName);
 
  var data;
  var label;
  
  if (id == "officialCompanydata"){
    data = wsData.getRange(2, column, wsData.getLastRow()-1, 2).getValues()
      .filter(function(item){
        var temp = item[1];
        if (temp == category) return item[0];});
    label = wsData.getRange(1,column).getValues()[0];
    data = data.map(function(k){return k[0];});
  }
  if (id == "internalCompanydata"){
    data = wsData.getRange(2, column, wsData.getLastRow()-1, 2).getValues()
      .filter(function(item){
        var temp = item[0];
        if (temp == category) return item[1];});
    label = wsData.getRange(1,column+1).getValues()[0];
    data = data.map(function(k){return k[1];});
  }
  data = unique(data);
  label = label[0];
  return [funcId, data, id, label, idlabel];
}

function getProviderData(funcId, ssName, id, column, idlabel, provider){
//ssName = 'Реестр договоров' 
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssName);

//id = 'officialCompanydata'
//column = 2
//idlabel = 'officialCompany'
//provider = 'Поставщик'

  var data;
  var label;

  if (id == "officialCompanydata"){
    data = wsData.getRange(2, column, wsData.getLastRow()-1, 6).getValues()
      .filter(function(item){
        var temp = item[5];
        if (temp == provider) return item[0];});
    label = wsData.getRange(1,column).getValues()[0];
    data = data.map(function(k){return k[0];});
  } 
  if (id == "internalCompanydata"){
    data = wsData.getRange(2, column, wsData.getLastRow()-1, 4).getValues()
      .filter(function(item){
        var temp = item[3];
        if (temp == provider) return item[0];});
    label = wsData.getRange(1,column).getValues()[0];
    data = data.map(function(k){return k[0];});
  }
  data = unique(data);
  label = label[0];
  return [funcId, data, id, label, idlabel];
}

function getDataTowns(funcId, ssName, id, column, idlabel){
  var category = getData(1, 'Тех. данные', 'categorydata', 1, 'category')[1];
//ssName = "Города";
//column = 1;
  var data = {};
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssName);
  var label = wsData.getRange(1,column).getValues()[0];
  var data = {0 : wsData.getRange(2,column, wsData.getLastRow() - 1).getValues().map(function(o){ return o[0];}).filter(function(o){ return o !== "" ;})};
  
  var allData = wsData.getRange(1, 1, 1, wsData.getLastColumn()).getValues()[0];
  var k = 1;
  for (var prop in category){
    allData.forEach(function (item, i){
      if (item == category[prop]){
        var triggers = wsData.getRange(2,i+1, wsData.getLastRow() - 1).getValues().map(function(o){ return o[0] ;});
        data[k] = triggers;
      }
    });
    k++;
  }
  label = label[0];
  return [funcId, data, id, label, idlabel, category];
}

function getDataWorks(funcId, ssName, id, column, idlabel){
  var category = getData(1, 'Тех. данные', 'categorydata', 1, 'category')[1];
//ssName = "Вид работ";
//column = 1;
  var data = {};
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName(ssName);
  var label = wsData.getRange(1,column).getValues()[0];
  var data = {0 : wsData.getRange(5,column, wsData.getLastRow() - 1).getValues().map(function(o){ return o[0];}).filter(function(o){ return o !== "";})};
  
  var allData = wsData.getRange(1, 1, 1, wsData.getLastColumn()).getValues()[0];
  var k = 1;
  for (var prop in category){
    allData.forEach(function (item, i){
      if (item == category[prop]){
        var triggers = wsData.getRange(5,i+1, wsData.getLastRow() - 1 - 3).getValues().map(function(o){ return o[0];});
        data[k] = triggers;
      }
    });
    k++;
  }
  label = label[0];
  return [funcId, data, id, label, idlabel, category];
}

function getUsers(funcId) {
  var user = PropertiesService.getUserProperties().getProperties();
  user = Session.getActiveUser().getEmail().toLowerCase();
  var wsDataTech = SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные");
  var adminAll = wsDataTech.getRange(2, 6, wsDataTech.getLastRow()).getValues();
  var emailAll = wsDataTech.getRange(2, 9, wsDataTech.getLastRow()).getValues();
  var admin;
  var id
  emailAll.forEach(function (item,i) {
    if (user == item){
      admin = adminAll[i][0];
      id = i;
    }
  });
  return [funcId, admin, id];
}  

function WriteData(funcId, obj, catID){
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Реестр договоров");
  
  var wsDataTech = SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные");
  var typeLabel = wsDataTech.getRange(1,5).getValues()[0][0];
  var typeClient = wsDataTech.getRange(2,5).getValues()[0][0];
  var categoryLabel = wsDataTech.getRange(1,1).getValues()[0][0];
  var companyLabel = wsDataTech.getRange(1,7).getValues()[0][0];

  var category, company;
  var type = true;
  
  var last = wsData.getLastRow()+1;
  var data = wsData.getRange(1,1,1,wsData.getLastColumn()).getValues()[0];
  data.forEach(function (item, i){
    for (var prop in obj.label) {
      if (obj.label[prop] == item){
        wsData.getRange(last,i+1).setValue(obj.value[prop]);
        if (item == categoryLabel){
          category = obj.value[prop];
        }
        if (item == companyLabel){
          company = obj.value[prop];
        }
        if (item == typeLabel){
          if (obj.value[prop] != typeClient){
            type = false;
        }}
      }
    }
  });
  
  if (type == true) {
    if (AccessToForms(catID, 1) == true){
      updateFormsCompany(companyLabel, category, catID); 
    }
  }
  
  return [funcId, true, company];
}

function WriteDataTowns(funcId, newTown, obj, list, catData){
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Города");
  var allData = wsData.getRange(1, 1, 1, wsData.getLastColumn()).getValues()[0];
  var row;
  var column;
  var last = wsData.getLastRow();
  
  if (newTown == false){  
    var towns = wsData.getRange(1,1,wsData.getLastRow(),1).getValues();

    towns.forEach(function (item, i) {
      if (item == obj[0]){
        row = i+1;
      }
    });
    allData.forEach(function (item, i){
      if (item == list[0]){
        column = i+1;
      }
    });
    list.forEach(function (item, i) {
      wsData.getRange(row, column+i).setValue(obj[1][i]);
    });
  } 
  if (newTown == true){
    wsData.getRange(last+1,1).setValue(obj[0]);

    allData.forEach(function (item, i){
      if (item == list[0]){
        column = i+1;
      }
    });
    list.forEach(function (item, i) {
      wsData.getRange(last+1, column+i).setValue(obj[1][i]);
    });
  }  
  SortTowns();
  LoadData(funcId, catData, 2);
  
  return[funcId, true, newTown, row, obj];
}

function WriteDataWorks(funcId, newWork, obj, list, catData){
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Вид работ");
  var allData = wsData.getRange(1, 1, 1, wsData.getLastColumn()).getValues()[0];
  var row;
  var column;
  var last = wsData.getLastRow();
  if (newWork == false){  
    var works = wsData.getRange(1,1,wsData.getLastRow(),1).getValues();

    works.forEach(function (item, i) {
      if (item == obj[0]){
        row = i+1;
      }
    });
    allData.forEach(function (item, i){
      if (item == list[0]){
        column = i+1;
      }
    });
    list.forEach(function (item, i) {
      wsData.getRange(row, column+i).setValue(obj[1][i]);
    });
  } 
  if (newWork == true){
    wsData.getRange(last+1,1).setValue(obj[0]);

    allData.forEach(function (item, i){
      if (item == list[0]){
        column = i+1;
      }
    });
    list.forEach(function (item, i) {
      wsData.getRange(last+1, column+i).setValue(obj[1][i]);
    });
  }  
//  SortWorks();
  LoadData(funcId, catData, 3);
  
  return[funcId, true, newWork, row, obj];
}

function ChooseCategory(){
  var html = HtmlService.createHtmlOutputFromFile('ChooseCategory.html');
  var ui = SpreadsheetApp.getUi();
  html.setWidth(610);
  html.setHeight(400); 
  ui.showModalDialog(html, 'Форма выбора категории клиента для прогрузки форм отчетов'); 
}

function getTargetLoad(funcId, key){
  var cache = CacheService.getScriptCache();
  var id = getUsers()[2];
  key = key+id;
  var cached = cache.get(key);
  if (cached != null) {
    return [funcId, cached]
  }
  return [funcId, 0];
}
function WriteCash(key, content, time){
  var cache = CacheService.getScriptCache();
  var user = getUsers()[2];
  key = key+user;
  cache.put(key, content, time); // cache for 10 minutes
}

function loadClients(){
  if (!CheckAccess()) return false;
  WriteCash("target", "1", 600);
  ChooseCategory();
}

function loadTowns(){
  if (!CheckAccess()) return false;
  WriteCash("target", "2", 600);
  ChooseCategory();
}

function loadWorks(){
  if (!CheckAccess()) return false;
  WriteCash("target", "3", 600);
  ChooseCategory();
}

function LoadData(funcId, obj, target){
  if (target == 1){
    var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Реестр договоров");
    var label = wsData.getRange(1, 4).getValues()[0];
    for (var prop in obj.value) {
      if (obj.value[prop] == true){
        AccessToForms(prop, 1);
        updateFormsCompany(label[0], obj.label[prop], prop);
      }
    }
  }
  if (target == 2){
    var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Города");
    var label = wsData.getRange(1, 1).getValues()[0];

    for (var prop in obj.value) {
      if (obj.value[prop] == true){
        AccessToForms(prop, 1);
        updateFormsTowns(label[0], obj.label[prop], prop);
      }
    }
  }
  if (target == 3){
    var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Вид работ");
    
    for (var prop in obj.value) {
      if (obj.value[prop] == true){
        AccessToForms(prop, 1);
        updateFormsWorks(obj.label[prop], prop);
      }
    }
  }
  return [funcId, true];
}

function FormDetect(category, id){
  var wsDataTech = SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные");
  
  var categoryAll = wsDataTech.getRange(2, 1, wsDataTech.getLastRow()).getValues();
  var formAll     = wsDataTech.getRange(2, 8, wsDataTech.getLastRow()).getValues();
  var form;
  
  categoryAll.forEach(function (item,i) {
    if (category == item){
      form = formAll[i][0];
    }
    if (id == i){
      form = formAll[i][0];
    }
  });
  return form;
}

function AccessToForms(id, access, form){
  var user = PropertiesService.getUserProperties().getProperties();
  user = Session.getActiveUser().getEmail();
//user = "adminmobdev@ezda.ru";
//  var id = 0;
//  var access = 1;
  if (user == "evgeny.lapkin@ezda.ru") return true;
  if (user == "mikhail.ryzhkin@ezda.ru") return true;
  if (user == "Active.Safety@com.ezda.ru") return true;
  
  //var allAccess = false;
  
  //if (access == 1){
  //  try {
  //    var editorsFrom = FormApp.openById(form).getEditors();
  //    editorsFrom.forEach(function (item){
  //      if (item == user && user != '') {
  //        allAccess = true;
  //      }
  //    });
  //  }
  //  catch (err){allAccess = false;}
  //}
  Logger.log("form ="+form)
  if (form == '' || form == undefined) {
    form = FormDetect('category', id);  
  } else {
   // Browser.msgBox("2 ="+form)
  }
  
  if (access == 1){
    var parameters = "?access="+access+"&user="+user+"&form="+form;
    var url = "https://script.google.com/macros/s/AKfycbxeUfRxc2AWi8ez45VUWZh2S-URcVkS2ngnMjlpD3o1Rh6tME4E/exec"+parameters;
    var res = UrlFetchApp.fetch(url);
    var response = res.getContentText();
    var response = res.getResponseCode();
    if (response == 200 )return true; 
    else return false;
  } 
  if (access == 0){
    var parameters = "?access="+access+"&user="+user+"&form="+form;
    var url = "https://script.google.com/macros/s/AKfycbxeUfRxc2AWi8ez45VUWZh2S-URcVkS2ngnMjlpD3o1Rh6tME4E/exec"+parameters;
    var res = UrlFetchApp.fetch(url);
    var response = res.getResponseCode();
    if (response == 200 )return true; 
    else return false;
  }
}

function updateFormsCompany(companyLabel, category, prop){
  var form = FormDetect(category, -1);
  form = FormApp.openById(form);
  
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Реестр договоров");
  var clientsFull = wsData.getRange(2, 3, wsData.getLastRow()-1, 2).getValues().filter(function(item){if (item[0] == category) return item[1];});
  var clients = clientsFull.map(function(k){return k[1].trim();});
  // clients = clients.trim();
  clients = unique(clients);
    // Browser.msgBox(clients)

  clients.sort();
  updateDropDownUsingTitle(companyLabel, clients, form, prop);
  
  if (category == "ГПН подрядчик"){
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(5,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(companyLabel, clients, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(5,8).getValues()[0][0]), 3);
    }
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(7,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(companyLabel, clients, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(7,8).getValues()[0][0]), 5);
    }
  }
  if (category == "ГПН ДО"){
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(6,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(companyLabel, clients, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(6,8).getValues()[0][0]), 4);
    }
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(8,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(companyLabel, clients, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(8,8).getValues()[0][0]), 6);
    }
  }
}

function updateFormsTowns(townLabel, category, prop){
  var form;
  var column;
  var form = FormDetect(category, -1);
  form = FormApp.openById(form);
  if (category == "ГПН подрядчик") {column = 4;}
  if (category == "ГПН ДО")        {column = 5;}
  if (category == "Прямые клиенты"){column = 6;}
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Города");

  var temp;
  var townsFull = wsData.getRange(2, 1, wsData.getLastRow()-1, 7).getValues()
      .filter(function(item){
        temp = item[column];
        temp = temp.replace(/\s+/g, ' ').trim();
        if (temp == "ДА" || temp == "Да" || temp == "да") return item[0];});
    //
  var towns = townsFull.map(function(k){return k[0];});
  towns = unique(towns);
  towns.sort();

  updateDropDownUsingTitle(townLabel, towns, form, prop);
  if (category == "ГПН подрядчик"){
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(5,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(townLabel, towns, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(5,8).getValues()[0][0]), 3);
    }
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(7,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(townLabel, towns, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(7,8).getValues()[0][0]), 5);
    }
  }
  if (category == "ГПН ДО"){
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(6,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(townLabel, towns, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(6,8).getValues()[0][0]), 4);
    }
    if (AccessToForms(100, 1, SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(8,8).getValues()[0][0]) == true){
      updateDropDownUsingTitle(townLabel, towns, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(8,8).getValues()[0][0]), 6);
    }
  }
//  if (category == "ГПН подрядчик"){
//    updateDropDownUsingTitle(townLabel, towns, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(5,8).getValues()[0][0]), 3);
//  }
//  if (category == "ГПН ДО"){
//    updateDropDownUsingTitle(townLabel, towns, FormApp.openById(SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные").getRange(6,8).getValues()[0][0]), 4);
//  }
}

function updateFormsWorks(category, prop){
  var column;
  var form = FormDetect(category, -1);
  form = FormApp.openById(form);
  if (category == "ГПН подрядчик") {column = 1;}
  if (category == "ГПН ДО")        {column = 2;}
  if (category == "Прямые клиенты"){column = 3;}
  var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Вид работ");

  var temp;
  var worksFull = wsData.getRange(2, 1, wsData.getLastRow()-1, 4).getValues()
      .filter(function(item){
        temp = item[column];
        temp = temp.trim();
        if (temp == "ДА" || temp == "Да" || temp == "да") return item[0];});
    //
  var works = worksFull.map(function(k){return k[0];});
  works = unique(works);
  //works.sort();
  if (category == "ГПН подрядчик"){
      label = wsData.getRange(2, 2).getValues()[0];
      updateDropDownUsingTitle(label[0], works, form, -1);
      label = wsData.getRange(3, 2).getValues()[0];
      updateDropDownUsingTitle(label[0], works, form, -1);
      label = wsData.getRange(4, 2).getValues()[0];
      updateDropDownUsingTitle(label[0], works, form, prop);}
  if (category == "ГПН ДО"){
      label = wsData.getRange(2, 3).getValues()[0];
      updateDropDownUsingTitle(label[0], works, form, -1);
      label = wsData.getRange(3, 3).getValues()[0];
      updateDropDownUsingTitle(label[0], works, form, -1);
      label = wsData.getRange(4, 3).getValues()[0];
      updateDropDownUsingTitle(label[0], works, form, prop);}
  if (category == "Прямые клиенты"){
      label = wsData.getRange(2, 4).getValues()[0];
      updateDropDownUsingTitle(label[0], works, form, prop);}  
 
}

function unique(arr)
{
    var n = arr.length, k = 0, arr2 = [];
    for (var i = 0; i < n; i++) 
     { var j = 0;
       while (j < k && arr2[j] !== arr[ i ]) j++;
       if (j == k) arr2[k++] = arr[ i ];
     }
    return arr2;
}

function updateDropDownUsingTitle(title, values, form, prop) {
  
  var items = form.getItems();
  var titles = items.map(function(item){
    return item.getTitle();
  });
  
  var pos = titles.indexOf(title);
  if(pos !== -1){
    var item = items[pos];
    var itemID = item.getId();
    updateDropdown(itemID, values, title, form, prop);
  }
}
  
function updateDropdown(id, values, title, form, prop) {
  var item = form.getItemById(id);
  if (item.getType() == 'MULTIPLE_CHOICE'){
    item.asMultipleChoiceItem().setChoiceValues(values);
  }
  if (item.getType() == 'LIST'){
    item.asListItem().setChoiceValues(values);
  }
  if (prop >= 0){
    AccessToForms(prop, 0);
  }
}

function SendMail(id, subject, obj, catData){
  var wsDataTech = SpreadsheetApp.openById(ssID).getSheetByName("Тех. данные");
  var emailAll = wsDataTech.getRange(2, 10, wsDataTech.getLastRow()).getValues();
  var emailaddress = '';
  emailAll.forEach(function(item){
    if (item != ""){
      if (emailaddress == ''){
        emailaddress = item[0];
      } else {
        emailaddress = emailaddress + ',' + item[0];
      }
    }  
  });
  var message = '';
  if (id == 1){
    for (var prop in obj.label) {
      message = message + obj.label[prop]+':<b> '+obj.value[prop]+' </b><br>';
    }
  }
  if (id == 2){
    message = '<b>'+obj[0]+' </b><br>';
    for (var prop in catData.label) {
      message = message + catData.label[prop]+':<b> '+obj[1][prop]+' </b><br>';
    }
  }
  //Browser.msgBox(emailaddress+"; "+subject+"; "+message);
 
  MailApp.sendEmail({
    to: emailaddress,
    subject: subject,
    htmlBody: message,
  }); 
}