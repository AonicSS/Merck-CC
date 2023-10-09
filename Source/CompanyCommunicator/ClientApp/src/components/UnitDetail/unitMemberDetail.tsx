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
  Add24Filled
} from "@fluentui/react-icons";
import { GetUnitMembersAction } from "../../actions";
import { deleteUnitMember, addUnitMember } from "../../apis/messageListApi";
import { useAppDispatch } from "../../store";

export const UnitMemberDetail = (unitMembers: any) => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });

  const dispatch = useAppDispatch();

  const addMember = async (id: number) => {
    try {
      await addUnitMember(id, { id: 4, name: "Jack Doe", email:"user4@email.com" });
      GetUnitMembersAction(dispatch, { id: 1 });
    } catch (error) {
      return error;
    }
  }

  const deleteMember = async (id: number) => {
    try {
      await deleteUnitMember(id, 1);
      GetUnitMembersAction(dispatch, { id: 1 });
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
          {unitMembers!.unitMembers!.map((item: any) => (
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
      <Button appearance="transparent" icon={<Add24Filled />} onClick={() => addMember(1)}>
        Add User
      </Button>
    </>
  );
};
