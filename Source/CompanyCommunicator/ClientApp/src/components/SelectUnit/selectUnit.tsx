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
import { GetUnitAction, GetUnitsAction, GetUserAction, GetUserUnitsAction, UpdateUserPermission } from '../../actions';
import { app } from '@microsoft/teams-js';
import { Header } from '../Shared/header';
import { IUser } from '../../models/user';
import { IUnit } from '../../models/unit';

interface ISelectUnit {
  theme: Theme;
}

const SelectUnit = (props: ISelectUnit) => {
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
  const navigate = useNavigate();
  const dispatch = useAppDispatch();

  const [userPrincipalName, setUserPrincipalName] = React.useState<string | undefined>(undefined);

  const getUpn = async () => {
    const ctx = await app.getContext();
    const upn = ctx.user?.userPrincipalName;
    setUserPrincipalName(upn);
  };

  React.useEffect(() => {
    void getUpn();
  }, []);

  const onSelectUnit = (props: string) => {
    GetUnitAction(dispatch, { id: props });
    navigate('/unitmessages');
  };

  const currentUser: IUser = useAppSelector((state: RootState) => state.messages).user.payload;
  const currentUserUnits: IUnit[] = useAppSelector((state: RootState) => state.messages).units.payload;
  const [unitFetched, setUnitFetched] = React.useState(false);

  // get the current user
  React.useEffect(() => {
    if (userPrincipalName) {
      GetUserAction(dispatch, { mail: userPrincipalName });
    }
  }, [dispatch, userPrincipalName]);

  // get the current users unit
  React.useEffect(() => {
    if (currentUser.id.length !== 0) {
      GetUserUnitsAction(dispatch, { id: currentUser.id });
      setUnitFetched(true);
    }
  }, [dispatch, currentUser]);

  // check if current user unit is empty after fetching
  React.useEffect(() => {
    if (currentUserUnits && currentUserUnits.length === 0 && unitFetched) {
      window.location.href = '/requestaccess';
    } else if (currentUserUnits && currentUserUnits.length !== 0) {
      const isInAdmin = currentUserUnits.some(unit => unit.name === 'Admin Unit');
      UpdateUserPermission(dispatch, isInAdmin);
    }
  }, [dispatch, currentUserUnits]);

  React.useEffect(() => {
    if (isAdmin) {
      const adminUnit = currentUserUnits.filter(unit => unit.name === 'Admin Unit');
      GetUnitsAction(dispatch);
      GetUnitAction(dispatch, { id: adminUnit[0].id });
      navigate('/messages');
    }
  }, [isAdmin]);

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
          {currentUserUnits?.map((item: any) => (
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
