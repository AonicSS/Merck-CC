// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { Spinner } from '@fluentui/react-components';
import { GetDraftMessagesAction, GetUnitDraftMessagesAction } from '../../actions';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import { DraftMessageDetail } from './draftMessageDetail';
import { IUnit } from '../../models/unit';

export const DraftMessages = () => {
  const { t } = useTranslation();
  const currentUnit: IUnit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const draftMessages = useAppSelector((state: RootState) => state.messages).draftMessages.payload;
  const loader = useAppSelector((state: RootState) => state.messages).isDraftMessagesFetchOn.payload;
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    if (!isAdmin) {
      GetUnitDraftMessagesAction(dispatch, { id: currentUnit.id });
    } else {
      GetDraftMessagesAction(dispatch)
    }
  }, [dispatch, currentUnit, isAdmin]);

  return (
    <>
      {loader && <Spinner labelPosition='below' label={t('fetching')} />}
      {draftMessages && draftMessages.length === 0 && !loader && <div>{t('EmptyDraftMessages')}</div>}
      {draftMessages && draftMessages.length > 0 && !loader && <DraftMessageDetail draftMessages={draftMessages} />}
    </>
  );
};
