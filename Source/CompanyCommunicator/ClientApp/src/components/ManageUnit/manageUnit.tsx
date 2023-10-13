import './manageUnit.scss';
import * as React from "react";
import { useTranslation } from "react-i18next";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Divider,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
} from "@fluentui/react-components";
import { PeopleAudience24Regular, MoreHorizontal24Filled, DeleteRegular, Pen24Regular } from "@fluentui/react-icons";
import { UnitMembers } from "../UnitMembers/unitMembers";
import { UnitGroups } from "../UnitGroups/unitGroups";
import { useAppDispatch, useAppSelector, RootState } from "../../store";
import { GetUnitAction } from "../../actions";
import { deleteUnit, updateUnit } from '../../apis/messageListApi';

export const ManageUnit = () => {
  const getUnit = () => {
    const search = window.location.search;
    const params = new URLSearchParams(search);
    const unit = params.get("unit");
    return unit ? parseInt(unit, 10) : 0;
  }

  const { t } = useTranslation();
  const dispatch = useAppDispatch();

  const unit = useAppSelector((state: RootState) => state.messages).unit.payload;

  React.useEffect(() => {
    if (unit && Object.keys(unit).length === 0) {
      GetUnitAction(dispatch, { id: getUnit() });
    }
  }, []);

  // select the current unit
  const currentUnit: any = unit;

  const renameUnit = async (id: number, name: string) => {
    try {
      await updateUnit(id, name);
      GetUnitAction(dispatch, { id: getUnit() });
    } catch (error) {
      return error;
    }
  }

  const removeUnit = async (id: number) => {
    try {
      await deleteUnit(id);
      window.location.href = `/selectUnit`;
    } catch (error) {
      return error;
    }
  }


  return (
    <>
      <div className="cc-unit">
        <div className="cc-unit-name">
          <PeopleAudience24Regular />
          <h2>{currentUnit?.name}</h2>
        </div>
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem key={'renameKey'} icon={<Pen24Regular />} onClick={() => renameUnit(currentUnit?.id, "new_name")}>
                Rename
              </MenuItem>
              <MenuItem key={'deleteKey'} icon={<DeleteRegular />} onClick={() => removeUnit(currentUnit?.id)}>
                {t('Delete')}
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
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
