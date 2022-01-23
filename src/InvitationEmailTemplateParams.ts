import { EmailTemplateParams } from './sheets-tools/EmailTemplate'

export class InvitationEmailTemplateParams implements EmailTemplateParams {
  sheetName = 'modèle courriel'
  conditions = []

  insertData (html: string, data: any): string {
    html = html.replaceAll('%PRENOM%', data.person.firstname)
    html = html.replaceAll('%NOM%', data.person.lastname)
    html = html.replaceAll('%EMAIL%', data.person.email)
    html = html.replaceAll('%OBJET%', data.meeting.subject)
    html = html.replaceAll('%DATE%', data.meeting.date.toLocaleDateString())
    html = html.replaceAll('%HEURE%', data.meeting.time)
    html = html.replaceAll('%LIEU%', data.meeting.venue)

    const htmlTemplate = HtmlService.createTemplateFromFile('dist/ListeSujets')
    const categories = [
      ...new Set(data.meeting.topics.map((x: { category: any }) => x.category))
    ] // use Set to get unique values
    htmlTemplate.data = categories.map(c => {
      return {
        category: c,
        topics: data.meeting.topics.filter(
          (t: { category: unknown }) => t.category === c
        )
      }
    })
    const htmlOutput = htmlTemplate.evaluate()
    const generated = htmlOutput.getContent()
    html = html.replaceAll('%LISTE_SUJETS%', generated)
    return html
  }

  evaluateCondition (_condition: string, _data: any): boolean {
    let conditionOk = true
    return conditionOk
  }
}
