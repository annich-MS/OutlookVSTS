enum WorkItemType {
    Bug,
    Task,
    UserStory,
}

export default WorkItemType;

export function typeToString(type: WorkItemType): string {
    switch (type) {
        case WorkItemType.Bug:
            return "Bug";
        case WorkItemType.Task:
            return "Task";
        case WorkItemType.UserStory:
            return "User Story";
        default:
            throw new Error("Unexepected work item type");
    }
}

export function typeFromString(type: string): WorkItemType {
    switch (type) {
        case "Bug":
            return WorkItemType.Bug;
        case "Task":
            return WorkItemType.Task;
        case "UserStory":
            return WorkItemType.UserStory;
        default:
            throw new Error(`Unexpected type ${type}`);
    }
}