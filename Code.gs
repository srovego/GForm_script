//1. Создание формы и размещение в требуемой папке
//2. Добавление элемента формы типа Список
// Тестовое изменение текстового поля для проверки синхронизации через github
const target_folder_id = '1506pJrf_J8UtjILlEUr9qjvjwcQjSkPt'
const ssID = "12IE4ePnIS4lGb8t8jZIsbuM03ASMfkfI5gCgHvXn8hk"
const ss_sheetName = "Sheet1"

function getWsHeaderData(gsheetId, sheetName){ 
  //Функция возвращает массив значений, соответсвующих заголовку 
  var wsData = SpreadsheetApp.openById(gsheetId).getSheetByName(sheetName)
  var headerData = wsData.getRange(1, 1, 1, wsData.getLastColumn()).getValues()[0].filter(function(o){return o !==""});;
  return headerData
}

function getRangeDataFromGSheet(gsheetId, sheetName, startCellRowNum = 1, startCellColNum = 1, rowsToRead = false, colsToRead = false){
  //Чтение требуемого диапазона данных с листа sheetName документа gsheetId, начиная с ячейки с координатами startCellRowNum, startCellCollNum
  // Если значения rowsToRead и colsToRead = False - возвращается массив, соответствующий всем доступным данным на листе. 
  var wsData = SpreadsheetApp.openById(gsheetId).getSheetByName(sheetName)
  wsDataColsNum = wsData.getLastColumn()
  wsDataRowsNum = wsData.getLastRow()

  if (rowsToRead !=false){
    wsRowsToRead = rowsToRead;
  }
  else{
    wsRowsToRead = wsDataRowsNum;
  }

  if (colsToRead !=false){
    wsColsToRead = colsToRead;
  }
  else{
    wsColsToRead = wsDataColsNum;
  }

  var values = wsData.getRange(startCellRowNum, startCellColNum, wsRowsToRead, wsColsToRead).getValues();
  Logger.log(values)
  return values
}

function addDropdownListToForm(formId, listTitle){
  //Функция добавляет новый выпадающий список (пустой) в указанную форму
  drop_item_name_1 = "Drop item 1"
  //form = DriveApp.getFileById(formId)
  form = FormApp.openById(formId)
  form.addListItem()
      .setTitle(listTitle)
      .setRequired(true)
}


function main(){
  //Тестовые запуски функций
  //form = createForm("A test form")
  //headerData = getWsHeaderData(gsheetId = ssID, sheetName = ss_sheetName)
  //getRangeDataFromGSheet(gsheetId = ssID, sheetName = ss_sheetName)
  
  //Фукнции, составляющие уже основную программу
  //Создаем пустую форму
  newFormId = createEmptyForm('An empty form', target_folder_id)
  
  //Считываем список вопросов
  questionsList = getWsHeaderData(ssID, ss_sheetName)
  
  //Создаем пустые выпадающие списки с соответствующими вопросами
  questionsList.forEach(function(question){
    addDropdownListToForm(formId = newFormId, listTitle = question)
  } );


  
  Logger.log(questionsList)
}


function main_temp(){
  var labels = wsData.getRange(1, 1, 1, wsData.getLastColumn()).getValues()[0];
  //Logger.log(labels)
  
  form = createForm("A test form")
  
  labels.forEach(function(label, i){
    var options = wsData
      .getRange(2, i + 1, wsData.getLastRow() - 1, 1)
      .getValues()
      .map(function(o){return o[0]})
      .filter(function(o){return o !==""});
    updateDropdownUsingTitle(form, label, options)
  });


}


function createEmptyForm(title, folderId = false){
//функция создает пустую форму и возвращает ее ID. При  необходимости - форма перемещается в папку folderId
  var item = title
  var form = FormApp.create(item)
    .setTitle(item);
  form.setDestination
  newFormId = form.getId()

  if (folderId != false){
    moveFileToFolder(newFormId, folderId)
  }
  return newFormId
}

function createForm(title) {
  var item = title
  var form = FormApp.create(item)
    .setTitle(item);
  form.setDestination
  
  //Move form to the specific folder
  var formId = form.getId()
  fld = DriveApp.getFolderById(target_folder_id)
  source = DriveApp.getFileById(formId);
  moveFileToFolder(fileId = formId, destinationFolderId = target_folder_id)

  //Adding items to the form
  item_name = "Text item 1"
  form.addTextItem()
      .setTitle(item_name)
      .setRequired(true);
  
  drop_item_name_1 = "Drop item 1"
  form.addListItem()
      .setTitle(drop_item_name_1)
      .setChoiceValues(["One", "Two"])
  
  drop_item_name_2 = "Drop item 2"
    form.addListItem()
        .setTitle(drop_item_name_2)
        .setChoiceValues(["One", "Two"])

  return form
}

function updateDropdownUsingTitle(form, title, values){
  var items = form.getItems()
  var titles = items.map(function(item){
    return item.getTitle();
  })
  var dropListPosition = titles.indexOf(drop_item_name)
  var dropListId = items[dropListPosition].asListItem().getId()
  updateDropdown(form, dropListId, values)
}


function updateDropdown(form, id, values){
  var item = form.getItemById(id)
  item.asListItem().setChoiceValues(values)

}

function moveFileToFolder(fileId, destinationFolderId){
  
  var file = DriveApp.getFileById(fileId);
  DriveApp.getFolderById(destinationFolderId).addFile(file);
  file
    .getParents()
    .next()
    .removeFile(file);


}

function getQuestionValues() {
  var ss= SpreadsheetApp.openById('1GeFzNR-UoFl9xbla8E1Ditsdr49UnrCmyGRr-6m9wNw');
  var questionSheet = ss.getSheetByName('Questions');
  var returnData = questionSheet.getDataRange().getValues();
  return returnData;
}

