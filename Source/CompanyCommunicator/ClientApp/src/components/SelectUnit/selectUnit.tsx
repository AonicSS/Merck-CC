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
import { HeaderContainer } from "../HeaderContainer/headerContainer";
import { GetUnitsAction } from "../../actions";

interface ISelectUnit {
  theme: Theme;
}

const SelectUnit = (props: ISelectUnit) => {
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });
  const { t } = useTranslation();
  
  function onSelectUnit(props: String) {
    window.location.href = `/unitmessages?unit=${props}`;
  }

  // TODO: Get current units of the user from api endpoint
  const units = useAppSelector((state: RootState) => state.messages).units.payload;
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    if (units && units.length === 0) {
      GetUnitsAction(dispatch);
    }
  }, []);

  const currentUserUnits = units;
  console.log(currentUserUnits);

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
                  onClick={() => onSelectUnit(item.name)}
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