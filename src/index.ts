import { InvitationEmailTemplateParams } from './InvitationEmailTemplateParams'
import { Meeting } from './models/Meeting'
import { Person } from './models/Person'
import { Parser } from './Parser'
import { EmailTemplate } from './sheets-tools/EmailTemplate'
import { getDriveId } from './sheets-tools/helpers'
import { Validation } from './sheets-tools/Validation'
import { ValidationHandler } from './sheets-tools/ValidationHandler'
;(global as any).onOpen = onOpen
;(global as any).onEdit = onEdit
;(global as any).onSendMeetingAgenda = onSendMeetingAgenda
;(global as any).onGenerateMeetingMinutes = onGenerateMeetingMinutes

function updateValidation (
  modifiedRange: GoogleAppsScript.Spreadsheet.Range | undefined = undefined
) {
  let validationHandler = ValidationHandler.getInstance()
  validationHandler.add(
    new Validation(
      'Sujets',
      'B2:B',
      '',
      '',
      false,
      Parser.getInstance()
        .meetings.map(x => x.id)
        .concat('à planifier')
    )
  )
  validationHandler.add(new Validation('Sujets', 'C2:C', 'Personnes', 'A2:A'))
  validationHandler.add(new Validation('Réunions', 'E2:E', 'Personnes', 'A2:A'))
  validationHandler.add(
    new Validation('Sujets', 'D2:D', 'Sujets', 'D2:D', true, ['Divers'])
  )
  validationHandler.add(new Validation('Tâches', 'A2:A', 'Personnes', 'A2:A'))
  validationHandler.add(
    new Validation('Tâches', 'D2:D', '', '', false, [
      'à faire',
      'fait',
      'en attente'
    ])
  )
  new EmailTemplate(new InvitationEmailTemplateParams()).addValidation()
  validationHandler.update(modifiedRange)
}

function onOpen () {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('Réunion')
    .addItem('Envoi ordre du jour', 'onSendMeetingAgenda')
    .addItem('Génération procès-verbal', 'onGenerateMeetingMinutes')
    .addToUi()
  updateValidation()
}

function onEdit (e: any) {
  updateValidation(e.range)
}

function getSelectedMeeting (): string {
  const sheetName = SpreadsheetApp.getActiveSheet().getSheetName()
  const currentRow = SpreadsheetApp.getCurrentCell().getRow()
  if (sheetName === 'Réunions' && currentRow !== 1) {
    const date = SpreadsheetApp.getActiveSheet()
      .getRange('A' + currentRow)
      .getValue()
    const time = SpreadsheetApp.getActiveSheet()
      .getRange('B' + currentRow)
      .getValue()
    const subject = SpreadsheetApp.getActiveSheet()
      .getRange('C' + currentRow)
      .getValue()
    const formattedDate = Utilities.formatDate(
      date,
      SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
      'dd.MM.yyyy'
    )
    return `${formattedDate} ${time} ${subject}`
  } else {
    return ''
  }
}

function onSendMeetingAgenda () {
  const meetingId = getSelectedMeeting()
  if (!meetingId) {
    SpreadsheetApp.getUi().alert(
      'Pour envoyer un ordre du jour, la ligne de la réunion concernée dans la feuille "Réunions" doit être sélectionnée.'
    )
    return
  }
  const meeting = Parser.getInstance().getMeeting(meetingId)
  if (!meeting) {
    SpreadsheetApp.getUi().alert(
      `Réunion avec identifiant "${meetingId}" introuvable!`
    )
    return
  }
  const template = new EmailTemplate(new InvitationEmailTemplateParams())

  // check if mail quota is sufficient
  if (MailApp.getRemainingDailyQuota() < meeting.attending.length) {
    SpreadsheetApp.getUi().alert(
      "Envoi impossible, quota d'envoi journalier dépassé!"
    )
    return
  }

  const concerned = ([] as Person[]).concat(
    meeting.attending,
    meeting.excused,
    meeting.missing
  )

  const uniqueConcerned = [
    ...new Map(concerned.map(item => [item['acronym'], item])).values()
  ]

  const data: { person: Person; meeting: Meeting }[] = uniqueConcerned.map(
    person => {
      return { person: person, meeting: meeting }
    }
  )
  let report: string[] = []
  for (const d of data) {
    const emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/
    if (emailPattern.test(d.person.email)) {
      const { subject, html } = template.constructHtml(d)
      const html2 = template.insertData(html, d)
      const subject2 = template.insertData(subject, d)
      try {
        MailApp.sendEmail({
          to: d.person.email,
          subject: subject2,
          htmlBody: html2
        })
        report.push(`Succès de l'envoi à ${d.person.name}.`)
      } catch (e) {
        report.push(`Echec de l'envoi à ${d.person.name} : ${e}.`)
      }
    } else {
      report.push(
        `Echec de l'envoi à ${d.person.name} : Email '${d.person.email}' invalide.`
      )
    }
  }
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      `<ul> ${report.map(x => `<li>${x}</li>`).join('')}</ul>`
    ),
    "Rapport d'envoi"
  )
}

function onGenerateMeetingMinutes () {
  const meetingId = getSelectedMeeting()
  if (!meetingId) {
    SpreadsheetApp.getUi().alert(
      'Pour générer un procès verbal, la ligne de la réunion concernée dans la feuille "Réunions" doit être sélectionnée.'
    )
    return
  }
  const meeting = Parser.getInstance().getMeeting(meetingId)
  if (!meeting) {
    SpreadsheetApp.getUi().alert(
      `Réunion avec identifiant "${meetingId}" introuvable!`
    )
    return
  }

  // delete current file in spreadsheet if existing
  // let sheet = SpreadsheetApp.getActiveSheet();
  // let currentId = sheet.getRange('H' + currentRow).getValue().match(/[-\w]{25,}(?!.*[-\w]{25,})/)
  // if (currentId) DriveApp.getFileById(currentId).setTrashed(true)

  // create new file from template
  const templateFileId = getDriveId(
    Parser.getInstance().params['modèle de procès verbal']
  )
  const outputFolderId = getDriveId(
    Parser.getInstance().params['dossier de génération']
  )
  let templateFile = DriveApp.getFileById(templateFileId)
  let destinationFolder = DriveApp.getFolderById(outputFolderId)
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
  docBody.replaceText('%HEURE_REUNION%', meeting.date.toLocaleTimeString())
  docBody.replaceText('%LIEU%', meeting.venue)
  docBody.replaceText('%DATE_REDACTION%', now.toLocaleDateString())
  docBody.replaceText('%AUTEUR%', meeting.author?.name ?? 'auteur')
  docHeader.replaceText('%OBJET%', meeting.subject)
  docHeader.replaceText('%DATE_REUNION%', meeting.date.toLocaleDateString())
  docHeader.replaceText('%HEURE_REUNION%', meeting.date.toLocaleTimeString())
  docHeader.replaceText('%LIEU%', meeting.venue)
  docHeader.replaceText('%DATE_REDACTION%', now.toLocaleDateString())
  docHeader.replaceText('%AUTEUR%', meeting.author?.name ?? 'auteur')
  docFooter.replaceText('%OBJET%', meeting.subject)
  docFooter.replaceText('%DATE_REUNION%', meeting.date.toLocaleDateString())
  docFooter.replaceText('%HEURE_REUNION%', meeting.date.toLocaleTimeString())
  docFooter.replaceText('%LIEU%', meeting.venue)
  docFooter.replaceText('%DATE_REDACTION%', now.toLocaleDateString())
  docFooter.replaceText('%AUTEUR%', meeting.author?.name ?? 'auteur')

  replacePlaceholderByList(
    '%PRESENTS%',
    meeting.attending,
    (x: { name: any; acronym: any }) => `${x.name} (${x.acronym})`
  )
  replacePlaceholderByList(
    '%EXCUSES%',
    meeting.excused,
    (x: { name: any; acronym: any }) => `${x.name} (${x.acronym})`
  )
  replacePlaceholderByList(
    '%ABSENTS%',
    meeting.missing,
    (x: { name: any; acronym: any }) => `${x.name} (${x.acronym})`
  )

  // add topics
  const topicsPlaceholderParagraphElement = docBody
    .findText('%SUJETS%') // find RangeElement containing text
    .getElement() // find corresponding TEXT Element
    .getParent() // find containing PARAGRAPH Element
  let topicIndex = docBody.getChildIndex(topicsPlaceholderParagraphElement)
  meeting.topics.forEach(topic => {
    docBody
      .insertParagraph(++topicIndex, topic.title)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    if (topic.category) {
      docBody.insertParagraph(++topicIndex, `Catégorie : ${topic.category}`)
    }
    if (topic.description) {
      docBody
        .insertParagraph(++topicIndex, 'Description')
        .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      docBody.insertParagraph(++topicIndex, topic.description)
    }
    if (topic.discussions) {
      docBody
        .insertParagraph(++topicIndex, 'Discussions')
        .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      docBody.insertParagraph(++topicIndex, topic.discussions)
    }
    if (topic.decisions) {
      docBody
        .insertParagraph(++topicIndex, 'Decisions')
        .setHeading(DocumentApp.ParagraphHeading.HEADING3)
      docBody.insertParagraph(++topicIndex, topic.decisions)
    }
  })
  topicsPlaceholderParagraphElement.removeFromParent()

  // add tasks
  const tasksTable = docBody
    .findText('%TÂCHES%') // find RangeElement containing text
    .getElement() // find corresponding Text Element
    .getParent() // find containing Paragraph Element
    .getParent() // find containing TableCell Element
    .getParent() // find containing TableRow Element
    .getParent() // find containing Table Element
    .asTable()
  // remove row with placeholder
  tasksTable.removeRow(tasksTable.getNumRows() - 1)
  Parser.getInstance()
    .tasks.filter(
      task => task.status === 'à faire' || task.status === 'en attente'
    )
    .forEach(task => {
      const row = tasksTable.appendTableRow()
      row.appendTableCell(task.assignee?.acronym ?? 'non défini')
      row.appendTableCell(task.dueDate.toLocaleDateString())
      row.appendTableCell(task.status)
      row.appendTableCell(task.description)
      if (task.dueDate < meeting.date) {
        let styleOverdue = {} as { [key: number]: string | boolean }
        styleOverdue[DocumentApp.Attribute.FOREGROUND_COLOR] = '#ff0000' // rouge
        styleOverdue[DocumentApp.Attribute.BOLD] = false
        row.setAttributes(styleOverdue)
      } else if (task.status === 'en attente') {
        let styleBlocked = {} as { [key: number]: string | boolean }
        styleBlocked[DocumentApp.Attribute.FOREGROUND_COLOR] = '#ff9900' // orange
        styleBlocked[DocumentApp.Attribute.BOLD] = false
        row.setAttributes(styleBlocked)
      } else {
        let styleDefault = {} as { [key: number]: string | boolean }
        styleDefault[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000' // black
        styleDefault[DocumentApp.Attribute.BOLD] = false
        row.setAttributes(styleDefault)
      }
    })

  // generate pdf
  fileToEdit.saveAndClose()
  let docblob = fileToEdit.getAs('application/pdf')
  let pdfFile: GoogleAppsScript.Drive.File
  let queryString = `title = '${fileName}.pdf' and mimeType = 'application/pdf'`
  if (destinationFolder.searchFiles(queryString).hasNext()) {
    pdfFile = destinationFolder.searchFiles(queryString).next()
    Drive.Files?.update(
      {
        title: `${fileName}.pdf`,
        mimeType: 'application/pdf'
      },
      pdfFile.getId(),
      docblob,
      {
        supportsAllDrives: true
      }
    )
  } else {
    pdfFile = destinationFolder.createFile(docblob)
  }

  // add link to new document in spreadsheet
  const documentUrl = `https://drive.google.com/file/d/${pdfFile.getId()}/view`
  SpreadsheetApp.getActiveSheet()
    .getRange('I' + SpreadsheetApp.getCurrentCell().getRow())
    .setValue(documentUrl)

  // delete doc file
  newFile.setTrashed(true)

  function replacePlaceholderByList (
    placeholder: string,
    list: any[],
    transform: Function
  ) {
    let element = docBody
      .findText(placeholder)
      .getElement()
      .getParent()
    let index = docBody.getChildIndex(element)
    element.removeFromParent()
    list.forEach(x =>
      docBody
        .insertListItem(index, transform(x))
        .setGlyphType(DocumentApp.GlyphType.BULLET)
    )
  }
}
