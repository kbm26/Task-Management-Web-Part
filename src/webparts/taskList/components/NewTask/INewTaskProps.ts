import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ITaskState } from "../TaskList";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

export interface INewTaskProps {
  setTasks: React.Dispatch<React.SetStateAction<ITaskState[]>>;
  tasks: ITaskState[];
  closeForm: () => void;
  existingItemId?: number;
  context: WebPartContext;
  isAdmin: boolean;
  user: ISiteUserInfo|undefined;
}
