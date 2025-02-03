import * as React from 'react';
import type { ITaskListProps } from './ITaskListProps';
import TaskItems from './TaskItems/TaskItems';
import NewTask from './NewTask/NewTask';
import { Panel, PrimaryButton } from '@fluentui/react';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { deleteTask, formatTaskList, getCurrentUser, getCurrentUserGroups, getItemsFromList, getItemsFromListForCurrentUser, isUserAdmin, updateTask } from '../services/Task.service';

export interface ITaskState {
  Title: string;
  Priority: string;
  DueDate: string;
  TaskId: number;
  Completed: boolean;
  Assignee: ISiteUserInfo;
}

const TaskList = (props: ITaskListProps): JSX.Element => {
  const [tasks, setTasks] = React.useState<ITaskState[]>([]);
  const [isOpen, setIsOpen] = React.useState<boolean>(false);
  const [selectedTask, setSelectedTask] = React.useState<ITaskState | undefined>(undefined);
  const [user, setUser] = React.useState<ISiteUserInfo | undefined>();
  const [isAdmin, setAdmin] = React.useState<boolean>(false);
  const { context } = props
  React.useEffect(() => {
    const fetchData = async (): Promise<void> => {
      try {
        const currentUser = await getCurrentUser();
        const groups = await getCurrentUserGroups(currentUser.Id);
        const admin = isUserAdmin(groups)
        setAdmin(admin)
        setUser(currentUser)
        const items = admin ? 
          await getItemsFromList() : 
          await getItemsFromListForCurrentUser(currentUser.Id)
        const mappedTasks = await formatTaskList(items)
        setTasks(mappedTasks);
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    };

    void fetchData();
  }, []);

  const onDeleteTask = async (taskId: number): Promise<void> => {
    try {
      await deleteTask(taskId)
      setTasks((prevTasks) => prevTasks.filter((task) => task.TaskId !== taskId));
    } catch (err) {
      console.error('Error deleting task:', err);
    }
  };

  const onUpdateTask = async (selectedTask: ITaskState): Promise<void> => {
    try {
      selectedTask.Completed = !selectedTask.Completed
      await updateTask(selectedTask)
      setTasks(tasks.map((task) =>
        task.TaskId === selectedTask.TaskId ? selectedTask : task
      ));
    } catch (err) {
      console.error('Error deleting task:', err);
    }
  }

  const openForm = (): void => {
    setSelectedTask(undefined);
    setIsOpen(true);
  };

  const handleEditTask = (task: ITaskState): void => {
    setSelectedTask(task);
    setIsOpen(true);
  };

  return (
    <>
      <>
        <h1>Todo Tasks</h1>
        <TaskItems isAdmin={isAdmin} tasks={tasks} onEditTask={handleEditTask} onUpdateTask={onUpdateTask} onDeleteTask={onDeleteTask} />
        {isAdmin && (
          <>
            <PrimaryButton text="Add new task" onClick={openForm} />
            <Panel
              key={selectedTask?.TaskId ?? "new-task"}
              headerText={selectedTask ? `Edit ${selectedTask.Title}` : "New Task"}
              isOpen={isOpen}
              closeButtonAriaLabel="Close"
              onDismiss={() => setIsOpen(false)}
            >
              <NewTask
                isAdmin={isAdmin}
                context={context}
                setTasks={setTasks}
                tasks={tasks}
                existingItemId={selectedTask?.TaskId}
                user={user}
                closeForm={() => setIsOpen(false)}
              />
            </Panel>
          </>
        )}
      </>
    </>
  );
};

export default TaskList;
/**
 * Update emails to be more dynamic
 * add changes to email.
 */