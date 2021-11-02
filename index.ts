function updateValidation(modifiedRange: GoogleAppsScript.Spreadsheet.Range = undefined) {
  let validationHandler = new ValidationHandler
  validationHandler.add(new Validation('Sujets', 'B2:B', 'Réunions', 'A2:A', false, ['à planifier']))
  validationHandler.add(new Validation('Sujets', 'C2:C', 'Personnes', 'A2:A'))
  validationHandler.add(new Validation('Réunions', 'D2:D', 'Personnes', 'A2:A'))
  validationHandler.add(new Validation('Sujets', 'D2:D', 'Sujets', 'D2:D', true))
  validationHandler.add(new Validation('Tâches', 'A2:A', 'Personnes', 'A2:A'))
  validationHandler.add(new Validation('Tâches', 'D2:D', '', '', false, ['à faire', 'fait', 'en attente']))
  validationHandler.update(modifiedRange)
}

function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Réunion')
      .addItem('Envoi ordre du jour', 'onSendMeetingAgenda')
      .addItem('Génération procès-verbal', 'onGenerateMeetingMinutes')
      .addToUi();
      updateValidation()
    }

function onEdit(e) {
  updateValidation(e.range)
}

function onSendMeetingAgenda() {
    
}

function onGenerateMeetingMinutes() {

  // detect meeting for which meeting minutes must be generated
  const sheetName = SpreadsheetApp.getActiveSheet().getSheetName()
  const currentRow = SpreadsheetApp.getCurrentCell().getRow()
  let date: Date
  let subject: string
  if(sheetName === 'Réunions') {
    date = SpreadsheetApp.getActiveSheet().getRange('A' + currentRow).getValue()
    subject = SpreadsheetApp.getActiveSheet().getRange('B' + currentRow).getValue()
  } else{
      SpreadsheetApp.getUi().alert('Pour envoyer un ordre du jour, la ligne de la réunion concernée dans la feuille "Réunions" doit être sélectionnée.')
      return
  }

  // get meeting data
  const parser = new Parser
  parser.parse()
  const meeting = parser.meetings.find(x => x.date.getTime() === date.getTime() && x.subject === subject)
  if (!meeting) {
    SpreadsheetApp.getUi().alert(`Réunion avec date "${date}" et sujet "${subject}" introuvable!`)
    return
  }

  // delete current file in spreadsheet if existing
  // let sheet = SpreadsheetApp.getActiveSheet();
  // let currentId = sheet.getRange('H' + currentRow).getValue().match(/[-\w]{25,}(?!.*[-\w]{25,})/)
  // if (currentId) DriveApp.getFileById(currentId).setTrashed(true)

  // create new file from template
  let templateFile = DriveApp.getFileById('1us4ErUoIChWcHvfM4tNDHfMhWb0yQrSw6ajV4gulu1c')
  let destinationFolder = DriveApp.getFolderById('1jWBay2PXXePEtcmBd6A_mYQZ-cqhZzDw')
  const fileName = `Réunion du ${meeting.date.toLocaleDateString()}`
  let newFile = templateFile.makeCopy(fileName, destinationFolder)
  var fileToEdit = DocumentApp.openById(newFile.getId())

  // replace placeholders in file
  let docBody = fileToEdit.getBody()
  let docHeader = fileToEdit.getHeader()
  let docFooter = fileToEdit.getFooter()
  let now = new Date()
  docBody.replaceText('%OBJET%', meeting.subject)
  docBody.replaceText('%DATE_REUNION%', meeting.date.toLocaleDateString())
  docBody.replaceText('%LIEU%', meeting.venue)
  docBody.replaceText('%DATE_REDACTION%', now.toLocaleDateString())
  docBody.replaceText('%AUTEUR%', meeting.author.name)
  docHeader.replaceText('%OBJET%', meeting.subject)
  docHeader.replaceText('%DATE_REUNION%', meeting.date.toLocaleDateString())
  docHeader.replaceText('%LIEU%', meeting.venue)
  docHeader.replaceText('%DATE_REDACTION%', now.toLocaleDateString())
  docHeader.replaceText('%AUTEUR%', meeting.author.name)
  docFooter.replaceText('%OBJET%', meeting.subject)
  docFooter.replaceText('%DATE_REUNION%', meeting.date.toLocaleDateString())
  docFooter.replaceText('%LIEU%', meeting.venue)
  docFooter.replaceText('%DATE_REDACTION%', now.toLocaleDateString())
  docFooter.replaceText('%AUTEUR%', meeting.author.name)
  
  replacePlaceholderByList('%PRESENTS%', meeting.attending, x => `${x.name} (${x.acronym})`)
  replacePlaceholderByList('%EXCUSES%', meeting.excused, x => `${x.name} (${x.acronym})`)
  replacePlaceholderByList('%ABSENTS%', meeting.missing, x => `${x.name} (${x.acronym})`)

  // add topics
  const topicsPlaceholderParagraphElement = docBody
    .findText('%SUJETS%') // find RangeElement containing text
    .getElement() // find corresponding TEXT Element
    .getParent() // find containing PARAGRAPH Element
  let index = docBody.getChildIndex(topicsPlaceholderParagraphElement)
  meeting.topics.forEach(topic => {
    docBody
      .insertParagraph(++index, topic.title)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    if(topic.category) {
      docBody.insertParagraph(++index, `Catégorie : ${topic.category}`)
    }
    if(topic.description) {
      docBody
        .insertParagraph(++index, 'Description')
        .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      docBody.insertParagraph(++index, topic.description)
    }
    if(topic.discussions) {
      docBody
        .insertParagraph(++index, 'Discussions')
        .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      docBody.insertParagraph(++index, topic.discussions)
    }
    if(topic.decisions) {
      docBody
        .insertParagraph(++index, 'Decisions')
        .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      docBody.insertParagraph(++index, topic.decisions)
    }
  })

  // generate pdf
  fileToEdit.saveAndClose()
  let docblob = fileToEdit.getAs('application/pdf')
  let pdfFile: GoogleAppsScript.Drive.File
  let queryString = `title = '${fileName}.pdf' and mimeType = 'application/pdf'`
  if(destinationFolder.searchFiles(queryString).hasNext()) {
    pdfFile = destinationFolder.searchFiles(queryString).next()
    Drive.Files.update({
      title: fileName, mimeType: 'application/pdf'
    }, pdfFile.getId(), docblob);
  } else {
    pdfFile = destinationFolder.createFile(docblob) 
  }

  // add link to new document in spreadsheet
  const documentUrl = `https://docs.google.com/document/d/${pdfFile.getId()}/edit`
  SpreadsheetApp.getActiveSheet().getRange('H' + currentRow).setValue(documentUrl)

  // delete doc file
  newFile.setTrashed(true)

  function replacePlaceholderByList(placeholder: string, list: any[], transform: Function) {
    let element = docBody
      .findText(placeholder)
      .getElement()
      .getParent();
    let index = docBody.getChildIndex(element);
    element.removeFromParent();
    list.forEach(x => docBody.insertListItem(index, transform(x)).setGlyphType(DocumentApp.GlyphType.BULLET));
  }
}