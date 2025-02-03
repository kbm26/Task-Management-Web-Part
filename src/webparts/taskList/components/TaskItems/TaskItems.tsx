import * as React from 'react';
import { ITaskState } from '../TaskList';
import { Checkbox, ContextualMenu, DefaultButton, DetailsList, Dialog, DialogFooter, DialogType, DirectionalHint, IColumn, IContextualMenuItem, PrimaryButton } from '@fluentui/react';
import { formatDate } from '../../services/Task.service';

interface TaskListProps {
  tasks: ITaskState[];
  onEditTask: (task: ITaskState) => void;
  isAdmin: boolean;
  onDeleteTask: (taskId: number) => void;
  onUpdateTask: (task: ITaskState) => void;
}

const TaskItems = (props: TaskListProps): JSX.Element => {
  const [menuTarget, setMenuTarget] = React.useState<HTMLElement | null>(null);
  const [isMenuVisible, setIsMenuVisible] = React.useState(false);
  const [selectedTask, setSelectedTask] = React.useState<ITaskState | null>(null);
  const [isDialogVisible, setIsDialogVisible] = React.useState(false);
  const [dialogMessage, setDialogMessage] = React.useState('');

  const { tasks, onEditTask, isAdmin, onDeleteTask, onUpdateTask } = props;
  
  const menuItems: IContextualMenuItem[] = [
    {
      key: 'edit',
      text: 'Edit',
      onClick: () => {
        if (selectedTask) {
          onEditTask(selectedTask);
        }
        setIsMenuVisible(false);
      },
    },
    {
      key: 'delete',
      text: 'Delete',
      onClick: () => {
        if (selectedTask) {
          setSelectedTask(selectedTask)
          setIsDialogVisible(true);
          setDialogMessage('Are you sure you want to delete this task?');
        }
        setIsMenuVisible(false);
      },
    },
  ];

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Title',
      fieldName: 'Title',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: 'priority',
      name: 'Priority',
      fieldName: 'Priority',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: 'dueDate',
      name: 'Due Date',
      fieldName: 'DueDate',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: ITaskState) => {
        return <span>{formatDate(item.DueDate)}</span>;
      },
    },
    {
      key: 'completed',
      name: 'Completed',
      fieldName: 'Completed',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: ITaskState) => (
        <Checkbox
          checked={item.Completed}
          disabled={isAdmin}
          onChange={async (_, checked) => {
            setIsDialogVisible(true);
            setDialogMessage(`Are you sure you want to mark this task as ${checked ? 'completed' : 'incomplete'}?`);
            setSelectedTask(item);
          }}
        />
      ),
    },
    {
      key: 'assignedTo',
      name: 'Assigned To',
      minWidth: 100,
      maxWidth: 150,
      onRender: (item: ITaskState) => (
        <span>{item.Assignee?.Title}</span>
      ),
    },
  ];

  if (isAdmin){
    columns.push({
      key: 'options',
      name: 'Options',
      minWidth: 100,
      maxWidth: 150,
      onRender: (item: ITaskState) => (
        <div
          style={{ cursor: 'pointer' }}
          onClick={(e: React.MouseEvent<HTMLDivElement>) => {
            setMenuTarget(e.currentTarget as HTMLElement);
            setSelectedTask(item);
            setIsMenuVisible(true);
          }}
        >
          ...
        </div>
      ),
    })
  }

  const handleDelete = async (): Promise<void> => {
    if (selectedTask) {
      try {
        onDeleteTask(selectedTask.TaskId);
      } catch (err) {
        console.error('Error deleting task:', err);
      }
    }
    setIsDialogVisible(false);
    setSelectedTask(null);
  };

  const handleConfirmComplete = async (): Promise<void> => {
    if (selectedTask) {
      try {
        onUpdateTask(selectedTask); 
      } catch (err) {
        console.error('Error updating task completion:', err);
      }
    }
    setIsDialogVisible(false);
    setSelectedTask(null);
  };

  const handleDialogClose = (): void => {
    setIsDialogVisible(false);
    setSelectedTask(null);
  };

  return (
    <div>
      <DetailsList
        items={tasks}
        columns={columns}
        selectionMode={0}
      />
      {isMenuVisible && isAdmin && (
        <ContextualMenu
          items={menuItems}
          target={menuTarget}
          directionalHint={DirectionalHint.bottomLeftEdge}
          onDismiss={() => {
            setIsMenuVisible(false);
          }}
        />
      )}

      <Dialog
        hidden={!isDialogVisible}
        onDismiss={handleDialogClose}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Confirm Action',
          subText: dialogMessage,
        }}
      >
        <DialogFooter>
          {(dialogMessage.indexOf('delete') !== -1)? (
            <PrimaryButton onClick={handleDelete} text="Delete" />
          ) : (
            <PrimaryButton onClick={handleConfirmComplete} text="Confirm" />
          )}
          <DefaultButton onClick={handleDialogClose} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default TaskItems;