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
  Combobox,
  Option,
  useId,
  ComboboxProps,
  Link,
} from "@fluentui/react-components";

import {
  DeleteRegular,
  PeopleAudience24Regular,
  MoreHorizontal24Filled,
  Add24Filled
} from "@fluentui/react-icons";
import { GetUsersAction, UpdateUnitAction } from "../../actions";
import { useAppDispatch, useAppSelector, RootState } from "../../store";
import { IUser } from "../../models/user";


export const UnitMemberDetail = () => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });

  const dispatch = useAppDispatch();

  // Get current units of the user from api endpoint
  const currentUnit: any = useAppSelector((state: RootState) => state.messages).unit.payload;
  const queryUsers = useAppSelector((state: RootState) => state.messages).users.payload;
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const unitUsers = currentUnit.users;
  const [filteredQueryUsers, setFilteredQueryUsers] = React.useState<IUser[]>([]);


  React.useEffect(() => {
    const filteredItems = queryUsers.filter(item => !unitUsers.some((user: any) => user.name === item.name));
    setFilteredQueryUsers(filteredItems);
  }, [queryUsers, currentUnit]);

  const addUser = async (item: IUser) => {
    if (item) {
      try {
        const updatedUnit = { ...currentUnit };
        updatedUnit.users = [...updatedUnit.users, item];
        UpdateUnitAction(dispatch, updatedUnit);
      } catch (error) {
        return error;
      }
    }
  }

  const deleteUser = async (userId: string) => {
    try {
      const updatedUnit = { ...currentUnit };
      updatedUnit.users = updatedUnit.users.filter((user: IUser) => user.id !== userId);
      UpdateUnitAction(dispatch, updatedUnit);
    } catch (error) {
      return error;
    }
  };

  const onSearchChange = (event: any) => {
    if (event?.target?.value) {
      const q = encodeURIComponent(event.target.value);
      GetUsersAction(dispatch, { query: q });
    }
  };


  const onSearchSelect: ComboboxProps['onOptionSelect'] = (event, data: any) => {
    const itemToAdd = {
      id: data.optionValue,
      name: data.optionText,
      mail: data.optionValue,
    }
    addUser(itemToAdd);
  };

  const comboId = useId("combo-default");

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
          {currentUnit?.users?.map((item: any) => (
            <TableRow key={'user' + item.id + 'key'}>
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
                  {isAdmin && <Menu>
                    <MenuTrigger disableButtonEnhancement>
                      <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        <MenuItem key={'deleteKey'} icon={<DeleteRegular />} onClick={() => deleteUser(item.id)}>
                          {t('Delete')}
                        </MenuItem>
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                  }
                </TableCellLayout>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
      <div style={{ display: "flex", alignItems: "center", gap: "15px", marginTop: "15px" }}>
        <Add24Filled />
        {isAdmin && <Combobox
          appearance='filled-darker'
          size='large'
          onOptionSelect={onSearchSelect}
          onChange={onSearchChange}
          aria-labelledby={comboId}
          placeholder="searchForUsers"
        >
          {filteredQueryUsers.map((opt) => (
            <Option text={opt.name} value={opt.id} key={opt.id}>
              {opt.name}
            </Option>
          ))}
          {filteredQueryUsers?.length === 0 ? (
            <Option key="no-results" text="">
              No results found
            </Option>
          ) : null}
        </Combobox>
        }
        {!isAdmin && <Link href="https://www.google.com" target="_blank">
          Add Users
        </Link>}
      </div>
    </>
  );
};