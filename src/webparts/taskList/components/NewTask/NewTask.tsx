import * as React from "react";
import {
  DatePicker,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  TextField,
} from "@fluentui/react";
import { INewTaskProps } from "./INewTaskProps";
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { addNewTask, editTask, getExistingTask, getUserById, getAssigneeIdsByEmail } from "../../services/Task.service";


const NewTask = (props: INewTaskProps): JSX.Element => {
  const { setTasks, closeForm, existingItemId, isAdmin } = props;
  const options: IDropdownOption[] = [
    { key: "Low", text: "Low" },
    { key: "High", text: "High" },
    { key: "Critical", text: "Critical" },
  ];

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory as any,
    spHttpClient: props.context.spHttpClient as any
  };

  const [title, setTitle] = React.useState<string>("");
  const [priority, setPriority] = React.useState<string>("Low");
  const [dueDate, setDueDate] = React.useState<Date>(new Date());
  const [completed, setCompleted] = React.useState<boolean>(false);
  const [assignees, setAssignees] = React.useState<string[]>([])

  React.useEffect(() => {
    const fetchExistingItem = async (): Promise<void> => {
      if (existingItemId) {
        try {
          const existingItem = await getExistingTask(existingItemId);
          const assignee = await getUserById(existingItem.AssigneeId);
          
          setTitle(existingItem.Title || "");
          setPriority(existingItem.Priority || "Low");
          setDueDate(existingItem.DueDate ? new Date(existingItem.DueDate) : new Date());
          setCompleted(existingItem.Completed || false);
          setAssignees([assignee.Email]);
        } catch (ex) {
          console.error("Error fetching the item:", ex);
        }
      } else {
        // No id to edit
      }
    };
  
    void fetchExistingItem();
  }, [existingItemId]);
  
  const saveForm = async (): Promise<void> => {
    if (!title || !dueDate) {
      alert("Please fill all required fields.");
      return;
    } else {
      try {
        const ids = await getAssigneeIdsByEmail(assignees);
        if (existingItemId) {
          await editTask(existingItemId, title,  priority, dueDate, ids)
          setTasks((prevTasks) =>
            prevTasks.map((task) =>
              task.TaskId === existingItemId
                ? { ...task, Title: title,  Priority: priority, DueDate: dueDate.toISOString(), Completed: completed }
                : task
            )
          );
        } else {
          const addedItem = await addNewTask(title,  priority, dueDate, ids);
          const user = await getUserById(ids[0].id);
          setTasks((prevTasks) => [
            ...prevTasks,
            {TaskId: addedItem.data.Id, Title: title, Priority: priority, DueDate: dueDate.toISOString(), Completed: completed, Assignee: user },
          ]);
        }
        closeForm();
      } catch (error) {
        console.error("Error saving task:", error);
        alert("Failed to save the task. Please try again.");
      }
    }
  };

  return (
    <>
      <TextField
        label="Title"
        placeholder="Title"
        value={title}
        onChange={(_e, newValue) => setTitle(newValue || "")}
      />
      <Dropdown
        label="Select Priority"
        placeholder="Priority"
        options={options}
        selectedKey={priority}
        onChange={(_e, option) => setPriority(option ? (option.key as string) : "Low")}
      />
      <DatePicker
        label="Due Date"
        placeholder="Date"
        value={dueDate}
        onSelectDate={(date) => setDueDate(date || new Date())}
      />
      {isAdmin && (
        <PeoplePicker
          context={peoplePickerContext as IPeoplePickerContext}
          titleText="Assignee"
          showtooltip={true}
          required={true}
          personSelectionLimit={1}
          onChange={(items) => setAssignees(items.map(item => item.secondaryText!))}
          principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
          defaultSelectedUsers={assignees}
        />
      )}
      <PrimaryButton text="Submit" onClick={saveForm} />
    </>
  );
};

export default NewTask;
