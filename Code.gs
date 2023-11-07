//1. Создание формы и размещение в требуемой папке
//2. Добавление элемента формы типа Список
// Тестовое изменение текстового поля для проверки синхронизации через github
const target_folder_id = '1506pJrf_J8UtjILlEUr9qjvjwcQjSkPt'
const ssID = "12IE4ePnIS4lGb8t8jZIsbuM03ASMfkfI5gCgHvXn8hk"
const ss_sheetName = "Sheet1"
const maxQuestionsOnOneSection = 3
const workFormId = "1bl1hixfq70ISrdYGvfWq3a7Ri-OWLrO072hZUbnNjv0"


function deleteListItems(formId){
  form = FormApp.openById(workFormId)
  listOfItemsToDelete = form.getItems(FormApp.ItemType.LIST)
  //listOfItemsToDelete.forEach(logTitleId)]
  listOfItemsToDelete.forEach(deleteElement)
}

function deleteElement(element){
  form.deleteItem(element)
}


function testSectionsCreation(){
  formId = "19FaTc9161gvIF8B9RgpQEb6PXL3vnrjouvPJemRrBsU"
  form = FormApp.openById(formId)

  const pageTwo = form.addPageBreakItem();

  addDropdownListToForm(formId = formId, listTitle = "Test list p2", points = 5)
  

  const pageThree = form.addPageBreakItem();
  addDropdownListToForm(formId = formId, listTitle = "Test list p3", points = 5)

  //Можно ли использовать идентификатор раздела..... - нет, возникает ошибка если использовать идентификатор раздела вместо идентификатора формы?
  sectionId = pageTwo.getId()
  sectionIndex = pageTwo.getIndex()
  //Logger.log("Индекс раздела 2: " + String(sectionIndex))

  pageTwo.setTitle('Page two');
  Logger.log("Идентификатор добавленной страницы 2: " + String(pageTwo.getId()) )
  Logger.log("Идентификатор добавленной страницы 3: " + String(pageThree.getId()) )

  pageThree.setTitle('Page three');

  allItems = form.getItems()
  allItems.forEach(logTitleId)
  //pageTwo.setGoToPage(pageThree);
  lastItem = allItems[allItems.length-1]

  theLastItemId = lastItem.getId()
  theLastItemIndex = lastItem.getIndex()
  form.moveItem(theLastItemIndex, allItems[1].getIndex())
  

}

function logTitleId(value){
  Logger.log(value.getTitle())
  Logger.log(value.getId())
  Logger.log(value.getIndex())
  Logger.log(value.getType())
}

function shuffleArray(array) {
  var i, j, temp;
  for (i = array.length - 1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}


function getWsHeaderData(gsheetId, sheetName){ 
  //Функция возвращает массив значений, соответсвующих заголовку 
  var wsData = SpreadsheetApp.openById(gsheetId).getSheetByName(sheetName)
  var headerData = wsData.getRange(1, 1, 1, wsData.getLastColumn()).getValues()[0].filter(function(o){return o !==""});;
  return headerData
}

function fillAnswerOptions(gsheetId, sheetName, formId, shuffle = false){
  var allData = getRangeDataFromGSheet(gsheetId, sheetName)
  //Logger.log(allData)
  var total_columns = allData[0].length
  
  form = FormApp.openById(formId)
  
  for (c=1; c<=total_columns; c++){
    Logger.log("C=" + String(c))
    question = allData[0][c-1]
    if (question != ""){
      //Если в заголовке, т.е. в тексте вопроса, что то указано - считываем все строки в данном столбце и заполняем форму
      answers = getRangeDataFromGSheet(gsheetId = ssID,
                                       sheetName = ss_sheetName,
                                       startCellRowNum = 2,
                                       startCellColNum = c,
                                       rowsToRead = false,
                                       colsToRead = 1).map(function(o){return o[0]}).filter( function(o){return o !==""} )
      correctAnswer = answers[0]
      Logger.log("Корректный ответ: " + String(correctAnswer))
      if (shuffle == true){
        answers = shuffleArray(answers)
      }
      updateDropdownUsingTitle(form = form, title = question, values = answers, correctAnswer = correctAnswer)
      Logger.log("Перемешанные ответы: " + String(answers))
      }
      else{
        Logger.log("No actions will be (empty question)")

      }
    }
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
  //Logger.log(values)
  return values
}

function addDropdownListToForm(formId, listTitle, points = false){
  //Функция добавляет новый выпадающий список (пустой) в указанную форму
  drop_item_name_1 = "Drop item 1"
  //form = DriveApp.getFileById(formId)
  form = FormApp.openById(formId)
  newItem = form.addListItem()
      .setTitle(listTitle)
      .setRequired(true)

  if (points > 0){
    newItem.setPoints(points)
  }
  return newItem
}


function main(){
  //Тестовые запуски функций
  //form = createForm("A test form")
  //headerData = getWsHeaderData(gsheetId = ssID, sheetName = ss_sheetName)
  //getRangeDataFromGSheet(gsheetId = ssID, sheetName = ss_sheetName)
  
  //Фукнции, составляющие уже основную программу
  //Создаем пустую форму
  newFormId = createEmptyForm('An empty form', target_folder_id, isQuiz = true)
  
  //Считываем список вопросов
  questionsList = getWsHeaderData(ssID, ss_sheetName)

  //Создаем необходимое кол-во разделов (секций)
  sectionsArray = createSections(formId = newFormId, questionsList = questionsList, questionsOnOneSection = maxQuestionsOnOneSection)

  //Создаем пустые выпадающие списки с соответствующими вопросами
  questionsArray = []
  questionsList.forEach(function(question){
    questionsArray.push(addDropdownListToForm(formId = newFormId, listTitle = question, points = 5) )
  } );
  
  //Заполняем опции ответов
  fillAnswerOptions(gsheetId = ssID, sheetName = ss_sheetName, formId = newFormId, shuffle = true)
  Logger.log(questionsList)

  //распределяем созданные вопросы по разделам
  distributeQuestionBySections(formId = newFormId, questionsArray, sectionsArray, questionsOnOneSection = maxQuestionsOnOneSection)
}

function distributeQuestionBySections(formId, questionsArray, sectionsArray, questionsOnOneSection){
  form = FormApp.openById(formId);
  questionsNumber = questionsArray.length;
  // sectionsNumber = sectionsArray.length;


  counter = 0;
  for (q = 0; q < questionsNumber; q++){
    
    section = Math.floor(q / questionsOnOneSection)
    sectionToMoveIndex = form.getItemById(sectionsArray[section]).getIndex() + 1
    form.moveItem(questionsArray[q].getIndex(), sectionToMoveIndex)
    Logger.log("Вопрос " + String(q) + " перемещается в секцию " + String(section))
  }


}


function createSections(formId, questionsList, questionsOnOneSection = 10){
  sectionsArray = []
  form = FormApp.openById(formId)

  var sectionsToCreate = Math.floor(questionsList.length / questionsOnOneSection)
  if (questionsList.length % questionsOnOneSection > 0){
    //Есть остаток от деления
    sectionsToCreate += 1;
  }

  Logger.log("Необходимо создать " + String(sectionsToCreate) + " разделов")

  for (s = 1; s <= sectionsToCreate; s++){
    var newSection = form.addPageBreakItem()
                          .setTitle("Группа вопросов " + String(s) + " из " + String(sectionsToCreate));
    sectionsArray.push(newSection.getId())                      
  }
  //Возвращаем массив идентификаторов разделов

  return sectionsArray
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


function createEmptyForm(title, folderId = false, isQuiz = false){
//функция создает пустую форму и возвращает ее ID. При  необходимости - форма перемещается в папку folderId
  var item = title
  var form = FormApp.create(item)
    .setTitle(item)
    .setPublishingSummary(true)
    .setConfirmationMessage("Форма заполнена. Результаты будут высланы на вашу почту.")
    .setCollectEmail(false)
    .setProgressBar(true)
    .setRequireLogin(true);
  form.setDestination
  newFormId = form.getId()
  
  if (isQuiz == true){
    form.setIsQuiz(true)
  }

  if (folderId != false){
    moveFileToFolder(newFormId, folderId)
  }
 
  return newFormId
}

function updateDropdownUsingTitle(form, title, values, correctAnswer = false){
  var items = form.getItems()
  var titles = items.map(function(item){
    return item.getTitle();
  })
  var dropListPosition = titles.indexOf(title)
  var dropListId = items[dropListPosition].asListItem().getId()
  updateDropdown(form, dropListId, values, correctAnswer)
  Logger.log("Обновление списка по имени. Заполняемые значения: " + String(values) + ". Корректное значение: " + String(correctAnswer))
}

function testSectionsCreation2(){
  //Inspired by https://www.youtube.com/watch?v=Adm7Ah-yyx8&t=312s

  formId = "19FaTc9161gvIF8B9RgpQEb6PXL3vnrjouvPJemRrBsU"
  form = FormApp.openById(formId)

  var sheets = ["Механика", "Электрика"]

  //Первый выбор - в какой раздел топать
  var classSelect = form.addMultipleChoiceItem();
  classSelect.setTitle("Выбери раздел")

  var classChoises = []
  for(var i = 0; i < sheets.length; i++){
    //Для каждого из элементов списка sheets
    var className = sheets[i];

    var classSection = form.addPageBreakItem()
        .setTitle(className)
        .setGoToPage(FormApp.PageNavigationType.SUBMIT);

    var students = ["Student 1", "Student 2", "Student 3"];
    var correctStatus = [true, false, false]
    
    //var studentSelect = form.addCheckboxItem()

    var studentSelect = form.addListItem()
        .setTitle(className + ' absent')
        .setHelpText("Укажи отсутствующих студентов");
    
    var studentChoices = [];
    for(var j = 0; j < students.length; j++){
      studentChoices.push(studentSelect.createChoice(students[j], correctStatus[j]));
    }
    //item.setChoices(
    //[item.createChoice(values[0], answersCorrectStatus[0]),


    studentSelect.setChoices(studentChoices)

    classChoises.push(classSelect.createChoice(className, classSection))


  }
  classSelect.setChoices(classChoises)

}

function updateDropdown(form, id, values, correctAnswer = false){
  var item = form.getItemById(id).asListItem()
  valuesNumber = values.length

  answersCorrectStatus = getCorrectStatusArray(values, correctAnswer)


  var itemChoices = [];
    for(var j = 0; j < values.length; j++){
      itemChoices.push(item.createChoice(values[j], answersCorrectStatus[j]));
    }
  item.setChoices(itemChoices)
  

}

function getCorrectStatusArray(values, correctAnswer = false){
  var result = []
  if (correctAnswer == false){
    for (i = 1; i < values.length; i++){
      result.push(false)
    }
  }
  else{
    for (i = 0; i < values.length; i++){
      if (values[i] == correctAnswer){
        result.push(true)
      }
      else{
        result.push(false)
      }
    }
  }
  return result
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