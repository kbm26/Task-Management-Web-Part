import {  FieldUserSelectionMode, IList, sp } from "@pnp/sp/presets/all";
import { IItemAddResult } from "@pnp/sp/items";
import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { ITaskState } from "../components/TaskList";

const LISTNAME = "Task List"

export const getListByTitle = async (): Promise<IList> => {
    return sp.web.lists.getByTitle(LISTNAME);
}

export const listExists = async (): Promise<boolean> => {
    const exists = await getListByTitle()
    try {
        console.log((await exists.get()))
        return true
    } catch (error) {
        console.log(error)
        return false
    }
}

export const createTaskList = async (): Promise<void> => {
    const { list } = await sp.web.lists.add(LISTNAME, "Tasks", 100, true, {
        Hidden: true,
    });
    console.log(list)

    await list.fields.addChoice("Priority", ["High", "Medium", "Low"]);
    await list.fields.addDateTime("DueDate");
    await list.fields.addNumber("TaskId");
    await list.fields.addBoolean("Completed");
    await list.fields.addUser("Assignee", FieldUserSelectionMode.PeopleOnly);
}

export const formatDate = (isoString: string): string => {
    const date = new Date(isoString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
};

export const getExistingTask = async (taskId: number): Promise<any> => {

    const list = await getListByTitle();
    const existingTask = list.items.getById(taskId)();
    return existingTask
}

export const getUserById = async (userId: number): Promise<ISiteUserInfo> => {
    return await sp.web.getUserById(userId)();
}

export const getAssigneeIdsByEmail = async (assigneeEmails: string[]): Promise<{ id: number, stringId: string }[]> => {
    const users = await sp.web.siteUsers();
    console.log(users, 'uisers')
    const ids: { id: number, stringId: string }[] = []
    for (const assignee of assigneeEmails) {
        const assigneeObject = await sp.web.siteUsers.getByEmail(assignee)();
        ids.push({ id: assigneeObject.Id, stringId: `${assigneeObject.Id}` })
    }
    return ids
}

export const editTask = async (existingTaskId: number, title: string, priority: string, dueDate: Date, userIds: { id: number, stringId: string }[]): Promise<void> => {
    const list = await getListByTitle();
    await list.items.getById(existingTaskId).update({
        Title: title,
        Priority: priority,
        DueDate: dueDate.toISOString(),
        Completed: false,
        AssigneeId: userIds[0].id,

    });
}

export const addNewTask = async (title: string, priority: string, dueDate: Date, userIds: { id: number, stringId: string }[]): Promise<IItemAddResult> => {
    const list = await getListByTitle();
    return await list.items.add({
        Title: title,
        Priority: priority,
        DueDate: dueDate.toISOString(),
        Completed: false,
        AssigneeId: userIds[0].id,
    });
}

export const getCurrentUser = async (): Promise<ISiteUserInfo> => {
    return await sp.web.currentUser()
}

export const getCurrentUserGroups = async (id: number): Promise<ISiteGroupInfo[]> => {
    return await sp.web.siteUsers.getById(id).groups()
}

export const isUserAdmin = (groups: ISiteGroupInfo[]): boolean => {
    return groups.filter((group) => group.Title === 'Task Admins').length > 0
}

export const getItemsFromList = async (): Promise<any[]> => {
    return (await getListByTitle()).items()
}

export const getItemsFromListForCurrentUser = async (userId: number): Promise<any[]> => {
    return (await getListByTitle()).items.filter(`AssigneeId eq '${userId}'`).get()
}

export const formatTaskList = async (tasks: any[]): Promise<ITaskState[]> => {
    return await Promise.all(tasks.map(async (task) => {
        const assignee = await getUserById(task.AssigneeId);
        return {
            TaskId: task.Id,
            Title: task.Title,
            Priority: task.Priority,
            DueDate: task.DueDate,
            Completed: task.Completed,
            Assignee: assignee,
        }
    }))
}

export const updateTask = async (selectedTask: ITaskState): Promise<void> => {
    const list = await getListByTitle();
    await list.items.getById(selectedTask.TaskId).update({
         Completed: selectedTask.Completed,
    });
}

export const deleteTask = async (taskId: number): Promise<void> => {
    const list = await getListByTitle();
    await list.items.getById(taskId).delete();
}
