class Parser {
    private _topics: Topic[] = []
    public get topics(): Topic[] {
        return this._topics
    }
    private _people: Person[] = []
    public get people(): Person[] {
        return this._people
    }
    private _meetings: Meeting[] = []
    public get meetings(): Meeting[] {
        return this._meetings
    }
    private _tasks: Task[] = []
    public get tasks(): Task[] {
        return this._tasks
    }

    parse(): void {
        // IMPORTANT : order is critical since parsing of meetings depends on people data, parsing of topics depends on people and meetings data
        this.parsePeople()
        this.parseMeetings()
        this.parseTopics()
        this.parseTasks()
    }

    private parseTasks() {
        this._tasks = []
        const tasksSheetValues = SpreadsheetApp.getActive().getSheetByName('Tâches').getDataRange().getValues()
        tasksSheetValues.shift() // shift removes first line that contains headings
        tasksSheetValues.forEach(row => {
            const assignee = this.people.find(x => x.acronym === row[0])
            const task = new Task(assignee, row[1], row[2], row[3])
            this.tasks.push(task)
            if (assignee)
                assignee.addTask(task)
        })
    }

    private parseTopics() {
        this._topics = []
        const topicsSheetValues = SpreadsheetApp.getActive().getSheetByName('Sujets').getDataRange().getValues()
        topicsSheetValues.shift() // shift removes first line that contains headings
        topicsSheetValues.forEach(row => {
            let meeting = this.meetings.find(x => x.date === row[0])
            const author = this.people.find(x => x.acronym === row[1])
            const topic = new Topic(row[0], meeting, author, row[3], row[4], row[5], row[6], row[7])
            this.topics.push(topic)
            if (meeting)
                meeting.addTopic(topic)
        })
    }

    private parseMeetings() {
        this._meetings = []
        const meetingsSheetValues = SpreadsheetApp.getActive().getSheetByName('Réunions').getDataRange().getValues()
        meetingsSheetValues.shift() // shift removes first line that contains headings
        meetingsSheetValues.forEach(row => {
            const author = this.people.find(x => x.acronym === row[3])
            const attendingAcronyms = row[4].trim().split(' ')
            const attending: Person[] = []
            attendingAcronyms.forEach(acronym => {
                const person = this.people.find(x => x.acronym === acronym)
                if (person)
                    attending.push(person)
            })
            const excusedAcronyms = row[5].trim().split(' ')
            const excused: Person[] = []
            excusedAcronyms.forEach(acronym => {
                const person = this.people.find(x => x.acronym === acronym)
                if (person)
                    excused.push(person)
            })
            const missingAcronyms = row[6].trim().split(' ')
            const missing: Person[] = []
            missingAcronyms.forEach(acronym => {
                const person = this.people.find(x => x.acronym === acronym)
                if (person)
                    missing.push(person)
            })
            this.meetings.push(new Meeting(row[0], row[1], row[2], author, attending, excused, missing))
        })
    }

    private parsePeople() {
        this._people = []
        const peopleSheetValues = SpreadsheetApp.getActive().getSheetByName('Personnes').getDataRange().getValues()
        peopleSheetValues.shift() // shift removes first line that contains headings
        peopleSheetValues.forEach(row => this.people.push(new Person(row[0], row[1], row[2])))
    }
}