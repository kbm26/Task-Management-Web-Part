import { sp } from "@pnp/sp";
import { IItemAddResult } from "@pnp/sp/items";
import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { ITaskState } from "../components/TaskList";

export const formatDate = (isoString: string): string => {
    const date = new Date(isoString);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
};

export const getEditableItem = async (itemId: number): Promise<any> => {
    const existingItem = await sp.web.lists
        .getByTitle("Todo Tasks")
        .items.getById(itemId)();
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

export const editTask = async (existingItemId: number, title: string, description: string, priority: string, dueDate: Date, userIds: { id: number, stringId: string }[]): Promise<void> => {
    await sp.web.lists.getByTitle("Todo Tasks").items.getById(existingItemId).update({
        Title: title,
        Description: description,
        Priority: priority,
        DueDate: dueDate.toISOString(),
        Completed: false,
        AssigneeId: userIds[0].id,
        AssigneeStringId: userIds[0].stringId
    });
}

export const addNewTask = async (title: string, description: string, priority: string, dueDate: Date, userIds: { id: number, stringId: string }[]): Promise<IItemAddResult> => {
    return await sp.web.lists.getByTitle("Todo Tasks").items.add({
    Title: title,
    Description: description,
    Priority: priority,
    DueDate: dueDate.toISOString(),
    Completed: false,
    AssigneeId: userIds[0].id,
    AssigneeStringId: userIds[0].stringId
    });      
}

export const getCurrentUser = async (): Promise<ISiteUserInfo> => {
    return await sp.web.currentUser()
}

export const getCurrentUserGroups = async (id: number): Promise<ISiteGroupInfo[]>  => {
   return await sp.web.siteUsers.getById(id).groups()
}

export const isUserAdmin = (groups: ISiteGroupInfo[]): boolean => {
   return groups.filter((group) => group.Title === 'Task Admins').length > 0
}

export const getItemsFromList = async(title: string): Promise<any[]> => {
    return await sp.web.lists.getByTitle(title).items()
}

export const getItemsFromListForCurrentUser = async(title: string, userId: number): Promise<any[]> => {
    return await sp.web.lists.getByTitle(title).items.filter(`AssigneeId eq '${userId}'`).get()
}

export const formatTaskList = async(items: any[]): Promise<ITaskState[]> => {
    return await Promise.all(items.map(async (item) => {
        const user = await sp.web.getUserById(item.AssigneeId)();
        return {
            TaskId: item.Id,
            Title: item.Title,
            Description: item.Description,
            Priority: item.Priority,
            DueDate: item.DueDate,
            Completed: item.Completed,
            User: user,
        }
    }))
}

export const updateTask = async(selectedTask: ITaskState): Promise<void> => {
    await sp.web.lists
        .getByTitle('Todo Tasks')
        .items.getById(selectedTask.TaskId)
        .update({
            Completed: selectedTask.Completed,
    });
}

export const deleteTask = async(taskId: number): Promise<void> => {
    await sp.web.lists
    .getByTitle("Todo Tasks")
    .items.getById(taskId)
    .delete();
}
