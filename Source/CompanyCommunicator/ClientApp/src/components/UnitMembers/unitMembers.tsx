// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation } from "react-i18next";
import { Spinner } from "@fluentui/react-components";
import { GetUnitMembersAction } from "../../actions";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { UnitMemberDetail } from "../UnitDetail/unitMemberDetail";

export const UnitMembers = () => {
  const { t } = useTranslation();
  const unit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const unitMembers = useAppSelector((state: RootState) => state.messages).unitMembers.payload;
  const loader = useAppSelector((state: RootState) => state.messages).isSentMessagesFetchOn.payload;
  const dispatch = useAppDispatch();

  const currentUnit: any = unit;

  React.useEffect(() => {
    if (unitMembers && unitMembers.length === 0) {
      GetUnitMembersAction(dispatch, { id: currentUnit?.id });
    }
  }, [currentUnit]);

  return (
    <>
      {loader && <Spinner labelPosition="below" label="Fetching..." />}
      {unitMembers && unitMembers.length === 0 && !loader && <div>No Members</div>}
      {unitMembers && unitMembers.length > 0 && !loader && <UnitMemberDetail />}
    </>
  );
};
