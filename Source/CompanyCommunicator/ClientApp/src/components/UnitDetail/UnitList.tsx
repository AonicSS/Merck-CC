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
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
} from '@fluentui/react-components';

import {
  DeleteRegular,
  PeopleAudience24Regular,
  MoreHorizontal24Filled,
  Pen24Regular,
} from '@fluentui/react-icons';
import * as microsoftTeams from '@microsoft/teams-js';
import { useAppDispatch, useAppSelector, RootState } from '../../store';
import { GetUnitsAction } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import { ROUTE_PARTS } from '../../routes';
import { deleteUnit } from '../../apis/messageListApi';
import { IUnit } from '../../models/unit';

export const UnitList = () => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
  const dispatch = useAppDispatch();

  const currentUserUnits = useAppSelector((state: RootState) => state.messages).units.payload;

  React.useEffect(() => {
    if (currentUserUnits && currentUserUnits.length === 0) {
      GetUnitsAction(dispatch);
    }
  }, []);

  const onManageUnit = (id: string) => {
    const unitUrl = getBaseUrl() + `/${ROUTE_PARTS.MANAGE_UNIT}/${id}/false`;
    const taskInfo: microsoftTeams.TaskInfo = {
      url: unitUrl,
      title: 'ManageUnit',
      height: microsoftTeams.TaskModuleDimension.Large,
      width: microsoftTeams.TaskModuleDimension.Large,
      fallbackUrl: unitUrl
    };

    const submitHandler = (result: any) => {
      if (result === null) {
        document.getElementById('manageUnitId')?.focus();
      } else {
        // TODO: handle manage unit and groups
        GetUnitsAction(dispatch);
      }
    };

    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
  };

  const removeUnit = async (id: string) => {
    try {
      await deleteUnit(id);
      GetUnitsAction(dispatch);
    } catch (error) {
      return error;
    }
  };

  return (
    <>
      <Table {...keyboardNavAttr} role='grid' aria-label='Select unit table with grid keyboard navigation'>
        <TableHeader>
          <TableRow>
            <TableHeaderCell key='title' style={{ width: '70%' }}>
              <b>Name</b>
            </TableHeaderCell>
            <TableHeaderCell key='userCount'>
              <b>Number of Admins</b>
            </TableHeaderCell>
            <TableHeaderCell key='actions' style={{ width: '50px' }}>
              <b>Actions</b>
            </TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {currentUserUnits?.map((unit: IUnit) => (
            <TableRow key={unit.id + 'key'}>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout
                  truncate
                  media={<PeopleAudience24Regular />}
                  style={{ cursor: 'pointer' }}
                  onClick={() => { onManageUnit(unit.id); }}
                >
                  {unit.name}
                </TableCellLayout>
              </TableCell>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout>{unit.size}</TableCellLayout>
              </TableCell>
              <TableCell role='gridcell' style={{ width: '50px' }}>
                <TableCellLayout>
                  <Menu>
                    <MenuTrigger disableButtonEnhancement>
                      <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        <MenuItem key={'editKey'} icon={<Pen24Regular />} onClick={() => { onManageUnit(unit.id); }}>
                          Edit
                        </MenuItem>
                        <Dialog>
                          <DialogTrigger>
                            <Button className='delete-dialog' icon={<DeleteRegular />}>{t('Delete')}</Button>
                          </DialogTrigger>
                          <DialogSurface>
                            <DialogBody>
                              <DialogTitle>Delete Unit?</DialogTitle>
                              <DialogContent>
                                Are you sure you want to delete this unit. This action cannot be undone.
                              </DialogContent>
                              <DialogActions>
                                <DialogTrigger disableButtonEnhancement>
                                  <Button appearance='secondary'>Close</Button>
                                </DialogTrigger>
                                <DialogTrigger>
                                  <Button onClick={() => { void removeUnit(unit.id); }} appearance='primary'>Delete</Button>
                                </DialogTrigger>
                              </DialogActions>
                            </DialogBody>
                          </DialogSurface>
                        </Dialog>
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                </TableCellLayout>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </>
  );
};
