import { Meeting } from './Meeting'
import { Person } from './Person'

export class Topic {
  constructor (
    readonly creationDate: Date,
    readonly meeting: Meeting | undefined,
    readonly author: Person | undefined,
    readonly category: string,
    readonly title: string,
    readonly description: string,
    readonly discussions: string,
    readonly decisions: string
  ) {}
}
