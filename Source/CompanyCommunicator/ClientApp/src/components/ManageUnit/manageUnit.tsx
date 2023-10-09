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
import { PeopleAudience24Regular } from "@fluentui/react-icons";
import { UnitMembers } from "../UnitMembers/unitMembers";
import { UnitGroups } from "../UnitGroups/unitGroups";

export const ManageUnit = () => {
  const getUnit = () => {
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const unit = params.get("unit");
    return unit;
  }

  return (
    <>
      <div className="cc-unit-name">
        <PeopleAudience24Regular />
        <h2>{getUnit()}</h2>
      </div>
      <Divider />
      <div>Organize, edit and view your unit and all the AD groups you can send a message to.
        If you would like to add a member to your unit or request access to a new AD Group,
        use the request button belwo</div>
      <Accordion defaultOpenItems={["1", "2"]} multiple collapsible>
        <AccordionItem value="1" key="draftMessagesKey">
          <AccordionHeader>Unit</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            <UnitMembers />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value="2" key="sentMessagesKey">
          <AccordionHeader>AD Groups</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            <UnitGroups />
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </>
  )
}
