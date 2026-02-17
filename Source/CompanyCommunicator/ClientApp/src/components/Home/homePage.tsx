// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import '../Shared/main.scss';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Theme, Body1Stronger } from '@fluentui/react-components';
import { AddCircle24Regular } from '@fluentui/react-icons';
import { app, dialog, DialogDimension, UrlDialogInfo } from '@microsoft/teams-js';
import { GetUnitsAction } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import { ROUTE_PARTS } from '../../routes';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import { DraftMessages } from '../DraftMessages/draftMessages';
import { ScheduledMessages } from '../ScheduledMessages/scheduledMessages';
import { SentMessages } from '../SentMessages/sentMessages';
import { Header } from '../Shared/header';
import { IUnit } from '../../models/unit';
import { UnitList } from '../UnitDetail/UnitList';

interface IHomePage {
  theme: Theme;
}

export const HomePage = (props: IHomePage) => {
  const currentUnit: IUnit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const url = getBaseUrl() + `/${ROUTE_PARTS.MANAGE_UNIT}/${currentUnit.id}/true`;
  const { t } = useTranslation();
  const dispatch = useAppDispatch();

  const onManageUnit = () => {
    const dialogInfo: UrlDialogInfo = {
      url,
      title: 'ManageUnit',
      size: { height: DialogDimension.Large, width: DialogDimension.Large },
      fallbackUrl: url,
    };

    const submitHandler: dialog.DialogSubmitHandler = (result: dialog.ISdkResponse) => {
      GetUnitsAction(dispatch);
    };

    // now open the dialog
    if (app.isInitialized()) {
      dialog.url.open(dialogInfo, submitHandler);
    }
  };

  return (
    <>
      <Header theme={props.theme} />
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '0px 15px' }}>
        <h2>{currentUnit.name}</h2>
        <Button
          className='cc-button'
          icon={<AddCircle24Regular />}
          appearance='primary'
          onClick={onManageUnit}
        >
          Create new unit
        </Button>
      </div>
      <Accordion defaultOpenItems={['1', '2', '3']} multiple collapsible>
        <AccordionItem value='1' key='unitsKey'>
          <AccordionHeader>List of Units</AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            <UnitList />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value='2' key='draftMessagesKey'>
          <AccordionHeader><Body1Stronger>{t('DraftMessagesSectionTitle')}</Body1Stronger></AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            <DraftMessages />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value='3' key='scheduledMessagesKey'>
          <AccordionHeader><Body1Stronger>{t('ScheduledMessagesSectionTitle')}</Body1Stronger></AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            <ScheduledMessages />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value='4' key='sentMessagesKey'>
          <AccordionHeader><Body1Stronger>{t('SentMessagesSectionTitle')}</Body1Stronger></AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            <SentMessages />
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </>
  );
};
