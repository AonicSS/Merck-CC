// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import {
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  TableRow,
  useArrowNavigationGroup,
  Theme,
} from '@fluentui/react-components';
import {
  PeopleAudience24Regular,
} from '@fluentui/react-icons';
import { useNavigate } from 'react-router-dom';
import { GetUnitAction, GetUnitsAction, GetUserUnitsAction, UpdateUserPermission } from '../../actions';
import { app } from '@microsoft/teams-js';
import { Header } from '../Shared/header';
import { IUnit } from '../../models/unit';
import { getUser, getUserUnits } from '../../apis/messageListApi';

interface ISelectUnit {
  theme: Theme;
}

const SelectUnit = (props: ISelectUnit) => {
  const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
  const navigate = useNavigate();
  const dispatch = useAppDispatch();

  const displayUnits: IUnit[] = useAppSelector((state: RootState) => state.messages).units.payload;

  const onSelectUnit = (props: string) => {
    GetUnitAction(dispatch, { id: props });
    setTimeout(() => {
      navigate('/unitmessages');
    }, 500);
  };

  React.useEffect(() => {
    let active = true;

    const loadUnits = async () => {
      const ctx = await app.getContext();
      const upn = ctx.user?.userPrincipalName;
      if (!upn || !active) return;

      const currentUser = await getUser(upn);
      if (!currentUser?.id || !active) return;

      const userUnits: IUnit[] = (await getUserUnits(currentUser.id)) || [];
      if (!active) return;

      if (userUnits.length === 0) {
        window.location.href = '/requestaccess';
        return;
      }

      const isUserAdmin = userUnits.some(unit => unit.name === 'Admin Unit');
      UpdateUserPermission(dispatch, isUserAdmin);

      if (isUserAdmin) {
        GetUnitsAction(dispatch);
      } else {
        GetUserUnitsAction(dispatch, { id: currentUser.id });
      }
    };

    void loadUnits();

    return () => { active = false; };
  }, [dispatch]);

  return (
    <>
      <Header theme={props.theme} />
      <h3>Select your unit</h3>
      <br />
      <Table {...keyboardNavAttr} role='grid' aria-label='Select unit table with grid keyboard navigation'>
        <TableHeader>
          <TableRow>
            <TableHeaderCell key='title'>
              <b>Name</b>
            </TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {displayUnits?.map((item: any) => (
            <TableRow key={item.id}>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout
                  truncate
                  media={<PeopleAudience24Regular />}
                  style={{ cursor: 'pointer' }}
                  onClick={() => { onSelectUnit(item.id); }}
                >
                  {item.name}
                </TableCellLayout>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
      <div style={{ padding: '15px 0px' }}>Note: you can switch units from your dashboard by clicking on your units name</div>
    </>
  );
};

export default SelectUnit;
