// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';
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
  Tooltip,
  Label
} from '@fluentui/react-components';

import {
  DeleteRegular,
  PeopleAudience24Regular,
  MoreHorizontal24Filled,
  Add24Filled,
  Info16Regular
} from '@fluentui/react-icons';
import { GetUsersAction, UpdateUnitAction } from '../../actions';
import { useAppDispatch, useAppSelector, RootState } from '../../store';
import { IUser } from '../../models/user';

export const UnitMemberDetail = () => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });

  const dispatch = useAppDispatch();

  // Get current units of the user from api endpoint
  const currentUnit: any = useAppSelector((state: RootState) => state.messages).unit.payload;
  const queryUsers = useAppSelector((state: RootState) => state.messages).users.payload;
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const unitUsers = currentUnit.users;
  const [filteredQueryUsers, setFilteredQueryUsers] = React.useState<IUser[]>([]);
  const [inputValue, setInputValue] = React.useState('');

  const [isHovered, setIsHovered] = React.useState(false);

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
  };

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
      setInputValue(event.target.value);
      const q = encodeURIComponent(event.target.value);
      if (q.length > 4) {
        GetUsersAction(dispatch, { query: q });
      }
    } else {
      setInputValue('');
    }
  };

  const onSearchSelect: ComboboxProps['onOptionSelect'] = (event, data: any) => {
    const itemToAdd = {
      id: data.optionValue,
      name: data.optionText,
      mail: data.optionValue,
    };
    /* prevent empty user addition */
    void (itemToAdd.id?.length && addUser(itemToAdd));
    setInputValue('');
  };

  const comboId = useId('combo-default');

  return (
    <>
      <Table {...keyboardNavAttr} role='grid' aria-label='Unit user table with grid keyboard navigation'>
        <TableHeader>
          <TableRow>
            <TableHeaderCell key='title'>
              <Label>User</Label>
              <Tooltip
                content={{
                  children: "Please provide at least 5 characters to begin searching for a user.",
                  
                }}
                relationship="label"
              >
                <Info16Regular
                  tabIndex={0}
                />
              </Tooltip>
            </TableHeaderCell>
            <TableHeaderCell key='actions' style={{ width: '50px' }}>
              <b>Actions</b>
            </TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {currentUnit?.users?.map((item: IUser) => (
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
                        <MenuItem key={'deleteKey'} icon={<DeleteRegular />} onClick={() => { void deleteUser(item.id); }}>
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
      <div style={{ display: 'flex', alignItems: 'center', gap: '15px', marginTop: '15px' }}>
        <Add24Filled />
        {isAdmin && <Combobox
          appearance='filled-darker'
          freeform
          size='large'
          onOptionSelect={onSearchSelect}
          onChange={onSearchChange}
          aria-labelledby={comboId}
          placeholder='Search for Users'
          value={inputValue}
        >
          {filteredQueryUsers.map((opt) => (
            <Option text={opt.name} value={opt.id} key={opt.id}>
              {opt.name}
            </Option>
          ))}
          {filteredQueryUsers?.length === 0
            ? (<Option key='no-results' text=''>No results found</Option>)
            : null}
        </Combobox>
        }
        {!isAdmin && <Link href='#'>
          Add Users
        </Link>}
      </div>
    </>
  );
};
