import { Person } from './Person'
import { Topic } from './Topic'

export class Meeting {
  readonly topics: Topic[] = []

  constructor (
    readonly date: Date,
    readonly time: string,
    readonly subject: string,
    readonly venue: string,
    readonly author: Person | undefined,
    readonly attending: Person[],
    readonly excused: Person[],
    readonly missing: Person[]
  ) {}

  public get id (): string {
    const formattedDate = Utilities.formatDate(
      this.date,
      SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
      'dd.MM.yyyy'
    )
    return `${formattedDate} ${this.time} ${this.subject}`
  }

  addTopic (topic: Topic): void {
    this.topics.push(topic)
  }
}
