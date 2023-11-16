// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./selectUnit.scss";
import React from "react";
import { useTranslation } from "react-i18next";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import {
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
  PeopleAudience24Regular,
} from "@fluentui/react-icons";
import {
  Theme,
} from "@fluentui/react-components";
import { useHistory } from 'react-router-dom';
import { HeaderContainer } from "../HeaderContainer/headerContainer";
import { GetUnitsAction, GetUserAction, GetUserUnitsAction, UpdateUserPermission } from "../../actions";
import * as microsoftTeams from '@microsoft/teams-js';

interface ISelectUnit {
  theme: Theme;
}

const SelectUnit = (props: ISelectUnit) => {
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });
  const { t } = useTranslation();
  const history = useHistory();
  const dispatch = useAppDispatch();

  const [userPrincipalName, setUserPrincipalName] = React.useState<string | undefined>(undefined);


  React.useEffect(() => {
    microsoftTeams.getContext(function (context) {
      setUserPrincipalName(context.userPrincipalName);
    });
  }, []);


  function onSelectUnit(props: string) {
    history.push(`/unitmessages?unit=${props}`);
  }

  const currentUser:any = useAppSelector((state: RootState) => state.messages).user.payload;
  const currentUserUnits = useAppSelector((state: RootState) => state.messages).units.payload;
  const [unitFetched, setUnitFetched] = React.useState(false);

  // get the current user
  React.useEffect(() => {
    if (userPrincipalName) {
      GetUserAction(dispatch, { mail: userPrincipalName });
    }
  }, [dispatch, userPrincipalName]);


  // get the current users unit
  React.useEffect(() => {
    if (Object.keys(currentUser).length !== 0) {
      GetUserUnitsAction(dispatch, { id: currentUser.id });
      setUnitFetched(true);
    }
  }, [dispatch, currentUser]);

  // check if current user unit is empty after fetching
  React.useEffect(() => {
    if (currentUserUnits && currentUserUnits.length === 0 && unitFetched) {
      window.location.href = `/requestaccess`;
    } else if (currentUserUnits && currentUserUnits.length !== 0) {
      const isInAdmin = currentUserUnits.some(unit => unit.name === "Admin Unit");
      UpdateUserPermission(dispatch, isInAdmin);
    }
  }, [dispatch, currentUserUnits]);

  React.useEffect(() => {
    if (isAdmin) {
      // if is admin, get all units
      GetUnitsAction(dispatch);
      history.push(`/messages`);
    }
  }, [dispatch, history, isAdmin]);

  return (
    <>
      <HeaderContainer theme={props.theme} />
      <h3>{t('SelectUnitText')}</h3>
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
                  onClick={() => onSelectUnit(item.id)}
                >
                  {item.name}
                </TableCellLayout>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
      <div>Note: you can switch units from your dashboard by clicking on your units name</div>
    </>
  );
};

export default SelectUnit;