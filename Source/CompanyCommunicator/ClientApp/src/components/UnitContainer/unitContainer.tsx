// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./unitContainer.scss";
import * as React from "react";
import { useTranslation } from "react-i18next";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Divider,
  Theme,
} from "@fluentui/react-components";
import { PeopleAudience24Regular, Settings24Filled, Status24Regular } from "@fluentui/react-icons";
import * as microsoftTeams from "@microsoft/teams-js";
import { GetDraftMessagesSilentAction, GetUnitAction } from "../../actions";
import { getBaseUrl } from "../../configVariables";
import { ROUTE_PARTS, ROUTE_QUERY_PARAMS } from "../../routes";
import { useAppDispatch, useAppSelector, RootState } from "../../store";
import { DraftMessages } from "../DraftMessages/draftMessages";
import { SentMessages } from "../SentMessages/sentMessages";
import { HeaderContainer } from "../HeaderContainer/headerContainer";

interface IUnitContainer {
  theme: Theme;
}


export const UnitContainer = (props: IUnitContainer) => {
  const getUnit = () => {
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const unit = params.get("unit");
    return unit ? unit : "0";
  }

  const { t } = useTranslation();
  const dispatch = useAppDispatch();

  const currentUnit: any = useAppSelector((state: RootState) => state.messages).unit.payload;

  React.useEffect(() => {
    if (currentUnit && Object.keys(currentUnit).length === 0) {
      GetUnitAction(dispatch, { id: getUnit() });
    }
  }, []);


  const messageUrl = getBaseUrl() + `/${ROUTE_PARTS.NEW_MESSAGE}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}&unit=${currentUnit?.id}`;
  const unitUrl = getBaseUrl() + `/${ROUTE_PARTS.MANAGE_UNIT}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}&unit=${currentUnit?.id}`;

  const onManageUnit = () => {
    let taskInfo: microsoftTeams.TaskInfo = {
      url: unitUrl,
      title: "ManageUnit",
      height: microsoftTeams.TaskModuleDimension.Large,
      width: microsoftTeams.TaskModuleDimension.Large,
      fallbackUrl: unitUrl
    };

    let submitHandler = (err: any, result: any) => {
      if (result === null) {
        document.getElementById("manageUnitId")?.focus();
      } else {
        microsoftTeams.tasks.submitTask();
      }
    };

    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
  }

  const onNewMessage = () => {
    let taskInfo: microsoftTeams.TaskInfo = {
      url: messageUrl,
      title: t("NewMessage"),
      height: microsoftTeams.TaskModuleDimension.Large,
      width: microsoftTeams.TaskModuleDimension.Large,
      fallbackUrl: messageUrl,
    };

    let submitHandler = (err: any, result: any) => {
      if (result === null) {
        document.getElementById("newMessageButtonId")?.focus();
      } else {
        GetDraftMessagesSilentAction(dispatch);
      }
    };

    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
  };

  return (
    <>
      <HeaderContainer theme={props.theme} />
      <Divider />
      <div className="cc-unit-container-header">
        <div className="cc-unit-name">
          <PeopleAudience24Regular />
          <h2>{currentUnit?.name}</h2>
        </div>
        <div>
          <div className="cc-new-message">
            <Button
              id="newMessageButtonId"
              icon={<Status24Regular />}
              appearance="primary"
              onClick={onNewMessage}
            >
              {t("NewMessage")}
            </Button>
          </div>
          <div className="cc-new-message">
            <Button
              id="manageUnitId"
              icon={<Settings24Filled />}
              appearance="primary"
              onClick={onManageUnit}
            >
              Manage unit & groups
            </Button>
          </div>
        </div>
      </div>
      <Accordion defaultOpenItems={["1", "2"]} multiple collapsible>
        <AccordionItem value="1" key="draftMessagesKey">
          <AccordionHeader>{t("DraftMessagesSectionTitle")}</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            {currentUnit.id && < DraftMessages />}
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value="2" key="sentMessagesKey">
          <AccordionHeader>{t("SentMessagesSectionTitle")}</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            {currentUnit.id && < SentMessages />}
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </>
  );
};
