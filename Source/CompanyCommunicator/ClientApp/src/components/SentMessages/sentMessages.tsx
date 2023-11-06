// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation } from "react-i18next";
import { Spinner } from "@fluentui/react-components";
import { GetSentMessagesAction, GetSentMessagesSilentAction, GetUnitSentMessagesAction } from "../../actions";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { SentMessageDetail } from "../MessageDetail/sentMessageDetail";
import * as CustomHooks from "../../useInterval";

export const SentMessages = () => {
  const { t } = useTranslation();
  const currentUnit: any = useAppSelector((state: RootState) => state.messages).unit.payload;
  const sentMessages = useAppSelector((state: RootState) => state.messages).sentMessages.payload;
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const loader = useAppSelector((state: RootState) => state.messages).isSentMessagesFetchOn.payload;
  const dispatch = useAppDispatch();
  const delay = 60000;

  React.useEffect(() => {
    if (!isAdmin) {
        GetUnitSentMessagesAction(dispatch, { id: currentUnit.id });
      } else {
        GetSentMessagesAction(dispatch)
      }
  }, [dispatch, currentUnit, isAdmin]);

  CustomHooks.useInterval(() => {
    GetSentMessagesSilentAction(dispatch);
  }, delay);

  return (
    <>
      {loader && <Spinner labelPosition="below" label="Fetching..." />}
      {sentMessages && sentMessages.length === 0 && !loader && <div>{t("EmptySentMessages")}</div>}
      {sentMessages && sentMessages.length > 0 && !loader && <SentMessageDetail sentMessages={sentMessages} />}
    </>
  );
};
