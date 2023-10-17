// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./mainContainer.scss";
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
import { Status24Regular, AddCircle24Regular } from "@fluentui/react-icons";
import * as microsoftTeams from "@microsoft/teams-js";
import { GetDraftMessagesSilentAction } from "../../actions";
import { getBaseUrl } from "../../configVariables";
import { ROUTE_PARTS, ROUTE_QUERY_PARAMS } from "../../routes";
import { useAppDispatch } from "../../store";
import { DraftMessages } from "../DraftMessages/draftMessages";
import { SentMessages } from "../SentMessages/sentMessages";
import { HeaderContainer } from "../HeaderContainer/headerContainer";
import { UnitList } from "../UnitList/unitList";

interface IMainContainer {
  theme: Theme;
}

export const MainContainer = (props: IMainContainer) => {
  const messageUrl = getBaseUrl() + `/${ROUTE_PARTS.NEW_MESSAGE}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}`;
  const unitUrl = getBaseUrl() + `/${ROUTE_PARTS.MANAGE_UNIT}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}&unit=new`;
  const { t } = useTranslation();
  const dispatch = useAppDispatch();

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
        //TODO: handle manage unit and groups
        console.log("Handle manage unit and groups");
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
          icon={<AddCircle24Regular />}
          appearance="primary"
          onClick={onManageUnit}
        >
          Create new Unit
        </Button>
      </div>
      <Accordion defaultOpenItems={["1", "2"]} multiple collapsible>
        <AccordionItem value="1" key="unitsKey">
          <AccordionHeader>List of Units</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            <UnitList />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value="2" key="draftMessagesKey">
          <AccordionHeader>{t("DraftMessagesSectionTitle")}</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            <DraftMessages />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value="3" key="sentMessagesKey">
          <AccordionHeader>{t("SentMessagesSectionTitle")}</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            <SentMessages />
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </>
  );
};
