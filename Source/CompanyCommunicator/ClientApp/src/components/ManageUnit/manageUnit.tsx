import './manageUnit.scss';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
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
  useId,
  Input,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
} from '@fluentui/react-components';
import * as microsoftTeams from '@microsoft/teams-js';
import { PeopleAudience24Regular, MoreHorizontal24Filled, DeleteRegular, Pen24Regular } from '@fluentui/react-icons';
import { useAppDispatch, useAppSelector, RootState } from '../../store';
import { GetUnitAction, GetUnitsAction, GetUserAction, GetUserUnitsAction, UpdateUnitAction, UpdateUserPermission } from '../../actions';
import { deleteUnit, updateUnit } from '../../apis/messageListApi';
import { useParams } from 'react-router-dom';
import { IUnit } from '../../models/unit';
import { UnitMemberDetail } from '../UnitDetail/UnitMemberDetail';
import { UnitGroupDetail } from '../UnitDetail/UnitGroupDetail';
import { app } from '@microsoft/teams-js';
import { IUser } from '../../models/user';

export const ManageUnit = () => {
  const { id, isNew } = useParams() as any;
  const { t } = useTranslation();
  const dispatch = useAppDispatch();

  const unit: IUnit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const currentUser: IUser = useAppSelector((state: RootState) => state.messages).user.payload;
  const currentUserUnits: IUnit[] = useAppSelector((state: RootState) => state.messages).units.payload;

  const [userPrincipalName, setUserPrincipalName] = React.useState<string | undefined>(undefined);
  const [currentUnit, setCurrentUnit]: any = React.useState(unit || []);
  const [isEditing, setIsEditing] = React.useState(false);

  const getUpn = async () => {
    const ctx = await app.getContext();
    const upn = ctx.user?.userPrincipalName;
    setUserPrincipalName(upn);
  };

  React.useEffect(() => {
    void getUpn();
    if (unit.id.length === 0) {
      GetUnitAction(dispatch, { id });
    }
  }, []);

  // get the current user
  React.useEffect(() => {
    if (userPrincipalName) {
      GetUserAction(dispatch, { mail: userPrincipalName });
    }
  }, [dispatch, userPrincipalName]);

  // get the current users unit
  React.useEffect(() => {
    if (currentUser.id.length !== 0) {
      GetUserUnitsAction(dispatch, { id: currentUser.id });
    }
  }, [dispatch, currentUser]);

  // check if current user unit is empty after fetching
  React.useEffect(() => {
    if (currentUserUnits && currentUserUnits.length !== 0) {
      const isInAdmin = currentUserUnits.some(unit => unit.name === 'Admin Unit');
      UpdateUserPermission(dispatch, isInAdmin);
    }
  }, [dispatch, currentUserUnits]);

  React.useEffect(() => {
    if (unit.name === 'Admin Unit' && isNew === 'true') {
      UpdateUnitAction(dispatch, {
        name: 'New_Unit',
        users: [],
        groups: [],
      });
    } else {
      setCurrentUnit(unit);
    }
  }, [unit]);

  const removeUnit = async () => {
    try {
      await deleteUnit(currentUnit.id);
      GetUnitsAction(dispatch);
      microsoftTeams.tasks.submitTask();
    } catch (error) {
      return error;
    }
  };

  const handleRenameClick = () => {
    setIsEditing(true);
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setCurrentUnit((currentUnit: any) => {
      const updatedUnit = { ...currentUnit, name: e.target.value };
      UpdateUnitAction(dispatch, updatedUnit);
      return updatedUnit; // Return the updated value to ensure it's reflected in the state
    });
  };

  const onNext = async (event: any) => {
    try {
      await updateUnit(currentUnit);
      microsoftTeams.tasks.submitTask();
    } catch (error) {
      return error;
    }
  };

  const inputId = useId('input');

  return (
    <>
      <div className='cc-unit'>
        <div className='cc-unit-name'>
          <PeopleAudience24Regular />
          {isEditing
            ? (<Input placeholder={currentUnit?.name} id={inputId} onChange={(event) => { handleInputChange(event); }} />)
            : (<h2>{currentUnit?.name}</h2>)}
        </div>
        {isAdmin && <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              {!isEditing && <MenuItem key={'renameKey'} icon={<Pen24Regular />} onClick={handleRenameClick}>
                Rename
              </MenuItem>}
              <Dialog>
                <DialogTrigger>
                  <Button className='delete-dialog' icon={<DeleteRegular />}>{t('Delete')}</Button>
                </DialogTrigger>
                <DialogSurface>
                  <DialogBody>
                    <DialogTitle>Delete Unit?</DialogTitle>
                    <DialogContent>
                      Are you sure you want to delete this unit. This action cannot be undone.
                    </DialogContent>
                    <DialogActions>
                      <DialogTrigger disableButtonEnhancement>
                        <Button appearance='secondary'>Close</Button>
                      </DialogTrigger>
                      <DialogTrigger>
                        <Button onClick={() => { void removeUnit(); }} appearance='primary'>Delete</Button>
                      </DialogTrigger>
                    </DialogActions>
                  </DialogBody>
                </DialogSurface>
              </Dialog>
            </MenuList>
          </MenuPopover>
        </Menu>}
      </div>
      <Divider />
      <div className='cc-unit'>Organize, edit and view your unit and all the AD groups you can send a message to.
        If you would like to add a member to your unit or request access to a new AD Group,
        use the request button below</div>
      <Accordion defaultOpenItems={['1', '2']} multiple collapsible>
        <AccordionItem value='1' key='unitMemberKey'>
          <AccordionHeader>Unit</AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            <UnitMemberDetail />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value='2' key='unitGroupKey'>
          <AccordionHeader>AD Groups</AccordionHeader>
          <AccordionPanel className='cc-accordion-panel'>
            <UnitGroupDetail />
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
      {isAdmin && <div className='fixed-footer'>
        <div className='footer-action-right'>
          { // eslint-disable-next-line @typescript-eslint/no-misused-promises
            <Button id='saveBtn' onClick={onNext} appearance='primary'>
              Save
            </Button>
          }
        </div>
      </div>
      }
    </>
  );
};
