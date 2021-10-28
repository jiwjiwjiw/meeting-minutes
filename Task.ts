class Task {
    constructor(
        readonly assignee: Person,
        readonly dueDate: Date,
        readonly description: string,
        readonly status: string
    ) {
        
    }
}