// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation } from "react-i18next";
import { Spinner } from "@fluentui/react-components";
import { GetUnitGroupsAction } from "../../actions";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { UnitGroupDetail } from "../UnitDetail/unitGroupDetail";

export const UnitGroups = () => {
  const { t } = useTranslation();
  const unit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const unitGroups = useAppSelector((state: RootState) => state.messages).unitGroups.payload;
  const loader = useAppSelector((state: RootState) => state.messages).isSentMessagesFetchOn.payload;
  const dispatch = useAppDispatch();

  const currentUnit: any = unit;

  React.useEffect(() => {
    if (unitGroups && unitGroups.length === 0) {
      GetUnitGroupsAction(dispatch, { id: currentUnit?.id });
    }
  }, [currentUnit]);


  return (
    <>
      {loader && <Spinner labelPosition="below" label="Fetching..." />}
      {unitGroups && unitGroups.length === 0 && !loader && <div>No Groups</div>}
      {unitGroups && unitGroups.length > 0 && !loader && <UnitGroupDetail />}
    </>
  );
};
