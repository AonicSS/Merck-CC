// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation } from "react-i18next";
import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  TableRow,
  useArrowNavigationGroup,
} from "@fluentui/react-components";
import {
  DeleteRegular,
  People24Regular,
  MoreHorizontal24Filled,
} from "@fluentui/react-icons";
import { GetUsersAction, UpdateUnitAction } from "../../actions";
import { deleteUnitMember, addUnitMember } from "../../apis/messageListApi";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { ComboBox } from "../ComboBox/comboBox";

interface MemberItem {
  id: number;
  name: string;
  email: string;
}

export const UnitMemberDetail = () => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });

  const dispatch = useAppDispatch();

  // Get current units of the user from api endpoint
  const currentUnit:any = useAppSelector((state: RootState) => state.messages).unit.payload;
  const users = useAppSelector((state: RootState) => state.messages).users.payload;
  const options = users?.filter((item: any) => !currentUnit.members?.some((elm: any) => elm.id === item.id));

  React.useEffect(() => {
    if (users && users.length === 0) {
      GetUsersAction(dispatch);
    }
  }, [dispatch, users, currentUnit]);



  const addMember = async (item: MemberItem) => {
    if (item) {
      try {
        const updatedUnit = await addUnitMember(currentUnit?.id, item);
        UpdateUnitAction(dispatch, updatedUnit);
      } catch (error) {
        return error;
      }
    }
  }

  const deleteMember = async (memberId: number) => {
    try {
      const updatedUnit = await deleteUnitMember(currentUnit?.id, memberId);
      UpdateUnitAction(dispatch, updatedUnit);
    } catch (error) {
      return error;
    }
  };

  return (
    <>
      <Table {...keyboardNavAttr} role='grid' aria-label='Unit user table with grid keyboard navigation'>
        <TableHeader>
          <TableRow>
            <TableHeaderCell key='title'>
              <b>User</b>
            </TableHeaderCell>
            <TableHeaderCell key='actions' style={{ width: '50px' }}>
              <b>Actions</b>
            </TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {currentUnit?.members?.map((item: any) => (
            <TableRow key={item.id + 'key'}>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout
                  truncate
                  media={<People24Regular />}
                >
                  {item.name}
                </TableCellLayout>
              </TableCell>
              <TableCell role='gridcell' style={{ width: '50px' }}>
                <TableCellLayout>
                  <Menu>
                    <MenuTrigger disableButtonEnhancement>
                      <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        <MenuItem key={'deleteKey'} icon={<DeleteRegular />} onClick={() => deleteMember(item.id)}>
                          {t('Delete')}
                        </MenuItem>
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                </TableCellLayout>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
      <ComboBox options={options} onSelect={addMember} placeholder="Add User" />
    </>
  );
};
