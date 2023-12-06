import { IGroup } from "./group";
import { IUser } from "./user";

export interface IUnit {
  id: string,
  name: string,
  size: number,
  groups: IGroup[],
  users: IUser[],
}