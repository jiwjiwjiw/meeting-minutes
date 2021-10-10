class Topic {
    constructor(
        readonly creationDate: Date,
        readonly meeting: Meeting,
        readonly author: Person,
        readonly title: string,
        readonly description: string,
        readonly discussions: string,
        readonly decisions: string,
        readonly tasks: Task[]) {
    }
}