class Person {
    private tasks: Task[] = []

    constructor(
        readonly acronym: string,
        readonly firstname: string,
        readonly lastname: string,
        readonly email: string
    ) {
    }

    public get name() : string {
        return `${this.firstname} ${this.lastname}`;
    }

    public addTask(task: Task): void {
        this.tasks.push(task)
    }
}