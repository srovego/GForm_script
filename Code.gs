//1. Создание формы и размещение в требуемой папке
//2. Добавление элемента формы типа Список
const target_folder_id = '1506pJrf_J8UtjILlEUr9qjvjwcQjSkPt'
const ssID = "12IE4ePnIS4lGb8t8jZIsbuM03ASMfkfI5gCgHvXn8hk"

var wsData = SpreadsheetApp.openById(ssID).getSheetByName("Sheet1")


function main(){
  


}


function main(){
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

