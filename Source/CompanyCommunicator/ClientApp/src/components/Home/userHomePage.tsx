// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import '../Shared/main.scss';
import './userHomePage.scss';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { useNavigate } from 'react-router-dom';
import { Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Theme, Body1Stronger } from '@fluentui/react-components';
import { Settings24Filled, Status24Regular, PeopleAudience24Regular, ChevronDown24Filled } from '@fluentui/react-icons';
import { app, dialog, DialogDimension, UrlDialogInfo } from '@microsoft/teams-js';
import { GetDraftMessagesSilentAction, GetUnitAction } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import { ROUTE_PARAMS, ROUTE_PARTS, ROUTE_QUERY_PARAMS } from '../../routes';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import { DraftMessages } from '../DraftMessages/draftMessages';
import { SentMessages } from '../SentMessages/sentMessages';
import { Header } from '../Shared/header';
import { ScheduledMessages } from '../ScheduledMessages/scheduledMessages';
import { IUnit } from '../../models/unit';

interface IHomePage {
  theme: Theme;
}

export const UserHomePage = (props: IHomePage) => {
  const currentUnit: IUnit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const currentUserUnits: IUnit[] = useAppSelector((state: RootState) => state.messages).units.payload;

  const messageUrl = getBaseUrl() + `/${ROUTE_PARTS.NEW_MESSAGE}/?${ROUTE_QUERY_PARAMS.LOCALE}={locale}&unitId=${currentUnit?.id}`;
  const unitUrl = getBaseUrl() + `/${ROUTE_PARTS.MANAGE_UNIT}/${currentUnit.id}/false`;

  const { t } = useTranslation();
  const dispatch = useAppDispatch();
  const navigate = useNavigate();

  const [isOpen, setIsOpen] = React.useState(false);

  const onManageUnit = () => {
    const dialogInfo: UrlDialogInfo = {
      url: unitUrl,
      title: "ManageUnit",
      size: { height: DialogDimension.Large, width: DialogDimension.Large },
      fallbackUrl: unitUrl,
    };

    const submitHandler: dialog.DialogSubmitHandler = (result: dialog.ISdkResponse) => {
      console.log("Handle here..");
    };

    // now open the dialog
    if (app.isInitialized()) {
      dialog.url.open(dialogInfo, submitHandler);
    }
  };

  const onNewMessage = () => {
    const dialogInfo: UrlDialogInfo = {
      url: messageUrl,
      title: t('NewMessage') ?? '',
      size: { height: DialogDimension.Large, width: DialogDimension.Large },
      fallbackUrl: messageUrl,
    };

    const submitHandler: dialog.DialogSubmitHandler = (result: dialog.ISdkResponse) => {
      GetDraftMessagesSilentAction(dispatch);
    };

    // now open the dialog
    if (app.isInitialized()) {
      dialog.url.open(dialogInfo, submitHandler);
    }
  };

  const handleToggleDropdown = () => {
    setIsOpen(!isOpen);
  };

  const handleSelectUnit = (unit: IUnit) => {
    GetUnitAction(dispatch, { id: unit.id });
    setIsOpen(false);
  };

  return (
    <>
      <Header theme={props.theme} />
      <div className="cc-user-homepage-header">
        <div className="cc-unit-dropdown">
          <div onClick={handleToggleDropdown} className="cc-unit-item">
            <PeopleAudience24Regular />
            <h2>{currentUnit.name}</h2>
            <ChevronDown24Filled />
          </div>
          {isOpen && (
            <ul style={{ listStyleType: 'none', margin: 0, padding: 0 }}>
              {currentUserUnits.filter(unit => unit.id !== currentUnit.id).map((unit) => (
                <li key={unit.id} onClick={() => handleSelectUnit(unit)} className="cc-unit-item">
                  <PeopleAudience24Regular />
                  <div>{unit.name}</div>
                </li>
              ))}
            </ul>
          )}
        </div>
        <div>
          <Button id='newMessageButtonId' className='cc-button' icon={<Status24Regular />} appearance='primary' onClick={onNewMessage}>
            {t('NewMessage')}
          </Button>
          {<Button id='manageUnitButtonId' className='cc-button' icon={<Settings24Filled />} appearance='primary' onClick={onManageUnit}>
            Manage Units
          </Button>
          }
        </div>
      </div>
      <Accordion defaultOpenItems={['1', '2', '3']} multiple collapsible>
        <AccordionItem value='1' key='draftMessagesKey'>
          <AccordionHeader><Body1Stronger>{t('DraftMessagesSectionTitle')}</Body1Stronger></AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            {currentUnit.id && <DraftMessages />}
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value='3' key='sentMessagesKey'>
          <AccordionHeader><Body1Stronger>{t('SentMessagesSectionTitle')}</Body1Stronger></AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            {currentUnit.id && <SentMessages />}
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </>
  );
};
