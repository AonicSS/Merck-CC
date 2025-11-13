// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { Spinner } from '@fluentui/react-components';
import { GetScheduledMessagesAction, GetUnitScheduledMessagesAction } from '../../actions';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import { ScheduledMessageDetail } from './scheduledMessageDetail';
import * as CustomHooks from '../../useInterval';
import { IUnit } from '../../models/unit';

export const ScheduledMessages = () => {
  const { t } = useTranslation();
  const currentUnit: IUnit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const scheduledMessages = useAppSelector((state: RootState) => state.messages).scheduledMessages.payload;
  const loader = useAppSelector((state: RootState) => state.messages).isScheduledMessagesFetchOn.payload;
  const dispatch = useAppDispatch();
  const delay = 60000;

  React.useEffect(() => {
    if (!isAdmin) {
      GetUnitScheduledMessagesAction(dispatch, { id: currentUnit.id });
    } else {
      GetScheduledMessagesAction(dispatch);
    }
  }, [dispatch, currentUnit, isAdmin]);

  CustomHooks.useInterval(() => {
    if (!isAdmin) {
      GetUnitScheduledMessagesAction(dispatch, { id: currentUnit.id });
    } else {
      GetScheduledMessagesAction(dispatch);
    }
  }, delay);

  console.log(scheduledMessages);

  return (
    <>
      {loader && <Spinner labelPosition='below' label='Fetching...' />}
      {scheduledMessages && scheduledMessages.length === 0 && !loader && <div>{t('EmptyScheduledMessages')}</div>}
      {scheduledMessages && scheduledMessages.length > 0 && !loader &&
        <ScheduledMessageDetail scheduledMessages={scheduledMessages} />}
    </>
  );
};
