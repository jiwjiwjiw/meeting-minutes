class Person {
    private tasks: Task[] = []

    constructor(
        readonly acronym: string,
        readonly name: string,
        readonly email: string
    ) {
    }
    
    addTask(task: Task): void {
        this.tasks.push(task)
    }
}