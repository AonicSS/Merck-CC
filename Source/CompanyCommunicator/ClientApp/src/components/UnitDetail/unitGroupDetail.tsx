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
} from "@fluentui/react-icons";
import { GetADGroupssAction, GetUnitGroupsAction } from "../../actions";
import { addUnitGroup, deleteUnitGroup } from "../../apis/messageListApi";
import { useAppDispatch, useAppSelector, RootState } from "../../store";
import { ComboBox } from "../ComboBox/comboBox";

interface GroupItem {
  id: number;
  name: string;
}

export const UnitGroupDetail = (unitGroups: any) => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });

  const dispatch = useAppDispatch();

  // Get current units of the user from api endpoint
  const unit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const groups = useAppSelector((state: RootState) => state.messages).adGroups.payload;
  const options = groups?.filter((item: any) => !unitGroups?.unitGroups?.some((elm: any) => elm.id === item.id));

  const currentUnit:any = unit;

  React.useEffect(() => {
    if (groups && groups.length === 0) {
      GetADGroupssAction(dispatch);
    }
  }, [dispatch, groups]);


  const addGroup = async (item: GroupItem) => {
    if (item) {
      try {
        await addUnitGroup(currentUnit?.id, item);
        GetUnitGroupsAction(dispatch, { id: currentUnit?.id });
      } catch (error) {
        return error;
      }
    }
  }

  const deleteGroup = async (groupId: number) => {
    try {
      await deleteUnitGroup(currentUnit?.id, groupId);
      GetUnitGroupsAction(dispatch, { id: currentUnit?.id });
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
      <ComboBox options={options} onSelect={addGroup} placeholder="Add AD Group" />
    </>
  );
};
