import { Person } from './Person'

export class Task {
  constructor (
    readonly assignee: Person | undefined,
    readonly dueDate: Date,
    readonly description: string,
    readonly status: string
  ) {}
}
