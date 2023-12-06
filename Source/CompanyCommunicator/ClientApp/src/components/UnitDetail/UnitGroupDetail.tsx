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
import { SearchGroupsAction, UpdateUnitAction } from "../../actions";
import { useAppDispatch, useAppSelector, RootState } from "../../store";
import { IGroup } from "../../models/group";


export const UnitGroupDetail = () => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });

  const dispatch = useAppDispatch();

  // Get current units of the user from api endpoint
  const currentUnit: any = useAppSelector((state: RootState) => state.messages).unit.payload;
  const queryGroups = useAppSelector((state: RootState) => state.messages).queryGroups.payload;
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const unitGroups = currentUnit.groups;
  const [filteredQueryGroups, setFilteredQueryGroups] = React.useState<IGroup[]>([]);


  React.useEffect(() => {
    const filteredItems = queryGroups.filter(item => !unitGroups.some((group: any) => group.name === item.name));
    setFilteredQueryGroups(filteredItems);
  }, [queryGroups, currentUnit]);

  const addGroup = async (item: IGroup) => {
    if (item) {
      try {
        const updatedUnit = { ...currentUnit }
        updatedUnit.groups = [...updatedUnit.groups, item];
        UpdateUnitAction(dispatch, updatedUnit);
      } catch (error) {
        return error;
      }
    }
  }

  const deleteGroup = async (groupId: string) => {
    try {
      const updatedUnit = { ...currentUnit };
      updatedUnit.groups = updatedUnit.groups.filter((group: any) => group.id !== groupId);
      UpdateUnitAction(dispatch, updatedUnit);
    } catch (error) {
      return error;
    }
  };

  const onSearchChange = (event: any) => {
    if (event?.target?.value) {
      const q = encodeURIComponent(event.target.value);
      SearchGroupsAction(dispatch, { query: q });
    }
  };

  const onSearchSelect: ComboboxProps['onOptionSelect'] = (event, data: any) => {
    const itemToAdd = {
      id: data.optionValue,
      name: data.optionText
    }
    addGroup(itemToAdd);
  };

  const comboId = useId("combo-default");

  return (
    <>
      <Table {...keyboardNavAttr} role='grid' aria-label='Unit user table with grid keyboard navigation'>
        <TableHeader>
          <TableRow>
            <TableHeaderCell key='title'>
              <b>Group</b>
            </TableHeaderCell>
            <TableHeaderCell key='actions' style={{ width: '50px' }}>
              <b>Actions</b>
            </TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {currentUnit?.groups?.map((item: any) => (
            <TableRow key={'group' + item.id + 'key'}>
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
                        <MenuItem key={'deleteKey'} icon={<DeleteRegular />} onClick={() => deleteGroup(item.id)}>
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
          placeholder={t('searchForGroups') ?? ''}
        >
          {filteredQueryGroups.map((opt) => (
            <Option text={opt.name} value={opt.id} key={opt.id}>
              {opt.name}
            </Option>
          ))}
          {filteredQueryGroups?.length === 0 ? (
            <Option key="no-results" text="">
              No results found
            </Option>
          ) : null}
        </Combobox>
        }
        {!isAdmin && <Link href="https://www.google.com" target="_blank">
          Add AD Groups
        </Link>}
      </div>
    </>
  );
};