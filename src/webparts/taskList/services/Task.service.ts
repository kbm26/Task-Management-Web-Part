import {  FieldUserSelectionMode, IList, sp } from "@pnp/sp/presets/all";
import { IItemAddResult } from "@pnp/sp/items";
import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { ITaskState } from "../components/TaskList";


export const getListByTitle = async (): Promise<IList> => {
    const bb = await sp.web.lists.getByTitle("Test Tasks").fields()
    console.log(bb)
    return sp.web.lists.getByTitle("Test Tasks");
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
    const { list } = await sp.web.lists.add("Test Tasks", "Tasks", 100, true, {
        Hidden: true,
    });
    console.log(list)

    await list.fields.addChoice("Priority", ["High", "Medium", "Low"]);
    await list.fields.addDateTime("DueDate");
    await list.fields.addNumber("TaskId");
    await list.fields.addBoolean("Completed");
    await list.fields.addUser("User", FieldUserSelectionMode.PeopleOnly);
}

export const formatDate = (isoString: string): string => {
    const date = new Date(isoString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
};

export const getExistingItem = async (itemId: number): Promise<any> => {
    const list = await getListByTitle();
    const existingItem = list.items.getById(itemId)();
    return existingItem
}

export const getUserById = async (userId: number): Promise<ISiteUserInfo> => {
    return await sp.web.getUserById(userId)();
}

export const getUserIdsByEmail = async (userEmails: string[]): Promise<{ id: number, stringId: string }[]> => {
    const ids: { id: number, stringId: string }[] = []
    for (const user of userEmails) {
        const userObject = await sp.web.siteUsers.getByEmail(user)();
        ids.push({ id: userObject.Id, stringId: `${userObject.Id}` })
    }
    return ids
}

export const editTask = async (existingItemId: number, title: string, priority: string, dueDate: Date, userIds: { id: number, stringId: string }[]): Promise<void> => {
    await (await getExistingItem(existingItemId)).update({
        Title: title,
        Priority: priority,
        DueDate: dueDate.toISOString(),
        Completed: false,
        UserId: userIds[0].id,

    });
}

export const addNewTask = async (title: string, priority: string, dueDate: Date, userIds: { id: number, stringId: string }[]): Promise<IItemAddResult> => {
    const list = await getListByTitle();
    return await list.items.add({
        Title: title,
        Priority: priority,
        DueDate: dueDate.toISOString(),
        Completed: false,
        UserId: userIds[0].id,
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
    return (await getListByTitle()).items.filter(`UserId eq '${userId}'`).get()
}

export const formatTaskList = async (items: any[]): Promise<ITaskState[]> => {
    console.log(items)
    return await Promise.all(items.map(async (item) => {
        const user = await getUserById(item.UserId);
        return {
            TaskId: item.Id,
            Title: item.Title,
            Priority: item.Priority,
            DueDate: item.DueDate,
            Completed: item.Completed,
            User: user,
        }
    }))
}

export const updateTask = async (selectedTask: ITaskState): Promise<void> => {
    await (await getExistingItem(selectedTask.TaskId)).update({
            Completed: selectedTask.Completed,
        });
}

export const deleteTask = async (taskId: number): Promise<void> => {
    await (await getExistingItem(taskId)).delete();
}
