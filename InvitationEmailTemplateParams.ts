class InvitationEmailTemplateParams implements EmailTemplateParams {
    sheetName = 'modèle courriel'
    conditions = []
    
    insertData(html: string, data: any): string {
        html = html.replaceAll('%PRENOM%', data.person.firstname)
        html = html.replaceAll('%NOM%', data.person.lastname)
        html = html.replaceAll('%EMAIL%', data.person.email)
        html = html.replaceAll('%OBJET%', data.meeting.subject)
        html = html.replaceAll('%DATE%', data.meeting.date.toLocaleDateString())
        html = html.replaceAll('%HEURE%', 'inconnu')
        html = html.replaceAll('%LIEU%', data.meeting.venue)
        html = html.replaceAll('%LISTE_SUJETS%', `<ol>${data.meeting.topics.map(x => `<li><h3>${x.title} (${x.author.name})</h3><p>${x.description}</p></li>`).join('')}</ol>`)
        return html
    }

    evaluateCondition(condition: string, data: any): boolean {
        let conditionOk = true
        return conditionOk
    }
}