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
  PeopleAudience24Regular,
  MoreHorizontal24Filled,
  Add24Filled
} from "@fluentui/react-icons";
import { GetUnitGroupsAction } from "../../actions";
import { addUnitGroup, deleteUnitGroup } from "../../apis/messageListApi";
import { useAppDispatch } from "../../store";

export const UnitGroupDetail = (unitGroups: any) => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });

  const dispatch = useAppDispatch();


  const addGroup = async (id: number) => {
    try {
      await addUnitGroup(id, { id: 1, name: "Test_Grp_4" });
      GetUnitGroupsAction(dispatch, { id: 1 });
    } catch (error) {
      return error;
    }
  }

  const deleteGroup = async (id: number) => {
    try {
      await deleteUnitGroup(id, 1);
      GetUnitGroupsAction(dispatch, { id: 1 });
    } catch (error) {
      return error;
    }
  };

  console.log(unitGroups);

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
          {unitGroups!.unitGroups!.map((item: any) => (
            <TableRow key={item.id + 'key'}>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout
                  truncate
                  media={<PeopleAudience24Regular />}
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
                        <MenuItem key={'deleteKey'} icon={<DeleteRegular />} onClick={() => deleteGroup(item.id)}>
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
      <Button appearance="transparent" icon={<Add24Filled />} onClick={() => addGroup(1)}>
        Add AD Group
      </Button>
    </>
  );
};
