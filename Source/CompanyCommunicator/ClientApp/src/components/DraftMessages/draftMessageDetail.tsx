// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';
import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Table,
  TableBody,
  TableCell,
  TableCellLayout,
  TableHeader,
  TableHeaderCell,
  TableRow,
  useArrowNavigationGroup,
  Body1Strong,
  Persona
} from '@fluentui/react-components';
import {
  DeleteRegular,
  DocumentCopyRegular,
  Chat20Regular,
  EditRegular,
  MoreHorizontal24Filled,
  OpenRegular,
  SendRegular,
} from '@fluentui/react-icons';
import { app, dialog, DialogDimension, UrlDialogInfo } from '@microsoft/teams-js';
import { GetDraftMessagesSilentAction, GetSentMessagesSilentAction, GetUnitDraftMessagesAction, GetUnitSentMessagesAction } from '../../actions';
import { deleteDraftNotification, duplicateDraftNotification, sendPreview } from '../../apis/messageListApi';
import { getBaseUrl } from '../../configVariables';
import { ROUTE_PARTS, ROUTE_QUERY_PARAMS } from '../../routes';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import { IUnit } from '../../models/unit';

export const DraftMessageDetail = (draftMessages: any) => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
  const isAdmin: boolean = useAppSelector((state: RootState) => state.messages).isAdmin.payload;
  const currentUnit: IUnit = useAppSelector((state: RootState) => state.messages).unit.payload;
  const [teamsTeamId, setTeamsTeamId] = React.useState('');
  const [teamsChannelId, setTeamsChannelId] = React.useState('');
  const [userPrincipalName, setUserPrincipalName] = React.useState<string | undefined>(undefined);
  const dispatch = useAppDispatch();
  const sendUrl = (id: string) => getBaseUrl() + `/${ROUTE_PARTS.SEND_CONFIRMATION}/${id}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}`;
  const editUrl = (id: string) => getBaseUrl() + `/${ROUTE_PARTS.NEW_MESSAGE}/${id}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}&unitId=${currentUnit?.id}`;
  const previewConfirmationUrl = () => getBaseUrl() + `/${ROUTE_PARTS.PREVIEW_MESSAGE_CONFIRMATION}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}`;

  React.useEffect(() => {
    if (app.isInitialized()) {
      void app.getContext().then((context: app.Context) => {
        setTeamsTeamId(context.team?.internalId ?? '');
        setTeamsChannelId(context.channel?.id ?? '');
        setUserPrincipalName(context.user?.userPrincipalName ?? '');
      });
    }
  }, []);

  const onOpenTaskModule = (url: string, title: string) => {
    const dialogInfo: UrlDialogInfo = {
      url,
      title,
      size: { height: DialogDimension.Large, width: DialogDimension.Large },
      fallbackUrl: url,
    };

    const submitHandler: dialog.DialogSubmitHandler = (result: dialog.ISdkResponse) => {
      if (!isAdmin) {
        GetUnitDraftMessagesAction(dispatch, { id: currentUnit.id });
        GetUnitSentMessagesAction(dispatch, { id: currentUnit.id });
      } else {
        GetDraftMessagesSilentAction(dispatch);
        GetSentMessagesSilentAction(dispatch);
      }
    };

    // now open the dialog
    dialog.url.open(dialogInfo, submitHandler);
  };

  const onPreviewMessageConfirmation = (url: string, title: string) => {
    const dialogInfo: UrlDialogInfo = {
      url,
      title,
      size: { height: DialogDimension.Small, width: DialogDimension.Small },
      fallbackUrl: url,
    };

    // now open the dialog
    dialog.url.open(dialogInfo);
  };

  const duplicateDraftMessage = async (id: number) => {
    try {
      await duplicateDraftNotification(id).then(() => {
        GetDraftMessagesSilentAction(dispatch);
      });
    } catch (error) {
      return error;
    }
  };

  const deleteDraftMessage = async (id: number) => {
    try {
      await deleteDraftNotification(id).then(() => {
        GetDraftMessagesSilentAction(dispatch);
      });
    } catch (error) {
      return error;
    }
  };

  const checkPreviewMessage = async (id: number) => {
    const payload = {
      draftNotificationId: id,
      teamsTeamId,
      teamsChannelId,
    };
    await sendPreview(payload)
      .then(() => {
        onPreviewMessageConfirmation(previewConfirmationUrl(), t('previewMessageTitle'));
        return true;
      })
      .catch((error) => {
        return error;
      });
  };

  // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
  const formatDate = (dateString: string): string => {
    const date = new Date(dateString); // Assuming dateString is in ISO 8601 format or a parsable date format

    // Format the date as desired, for example:
    const formattedDate = date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
    });

    return formattedDate;
  };

  return (
    <Table {...keyboardNavAttr} role='grid' aria-label={t('draftMessagesGridNavigation') ?? ''}>
      <TableHeader>
        <TableRow>
          <TableHeaderCell key='title'>
            <Body1Strong>{t('TitleText')}</Body1Strong>
          </TableHeaderCell>
          <TableHeaderCell key='groupName'>
            <b>Send to Group</b>
          </TableHeaderCell>
          <TableHeaderCell key='createdBy'>
            <b>{t('CreatedBy')}</b>
          </TableHeaderCell>
          <TableHeaderCell key='createdOn'>
            <b>Created On</b>
          </TableHeaderCell>
          <TableHeaderCell key='actions' style={{ width: '50px' }}>
            <Body1Strong>{t('actions')}</Body1Strong>
          </TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {draftMessages.draftMessages?.map((item: any) => (
          // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
          <TableRow key={`${item.id}key`}>
            <TableCell tabIndex={0} role='gridcell'>
              <TableCellLayout
                truncate
                media={<Chat20Regular />}
                style={{ cursor: 'pointer' }}
                onClick={() => {
                  onOpenTaskModule(editUrl(item.id), t('EditMessage'));
                }}
              >
                <Body1Strong style={{ whiteSpace: 'nowrap' }}>{item.title}</Body1Strong>
              </TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role='gridcell'>
              <TableCellLayout truncate title={item.groupNames}>
                {// eslint-disable-next-line @typescript-eslint/restrict-template-expressions
                  <div>{item.groupNames ?? [0]}</div>
                 }
              </TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role='gridcell'>
              <TableCellLayout truncate title={item.createdBy}>
                <Persona size='extra-small' textAlignment='center' name={item.createdBy} secondaryText={'Member'} avatar={{ color: 'colorful' }} />
              </TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role='gridcell'>
              <TableCellLayout truncate title={item.createdDateTime}>
                <div>{formatDate(item.createdDateTime)}</div>
              </TableCellLayout>
            </TableCell>
            <TableCell role='gridcell' style={{ width: '50px' }}>
              <TableCellLayout style={{ float: 'right' }}>
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <Button aria-label='Actions menu' icon={<MoreHorizontal24Filled />} />
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      {(isAdmin || userPrincipalName !== item.createdBy) && <MenuItem
                        icon={<SendRegular />}
                        key={'sendConfirmationKey'}
                        onClick={() => {
                          onOpenTaskModule(sendUrl(item.id), t('SendConfirmation'));
                        }}
                      >
                        {t('Send')}
                      </MenuItem>}
                      {
                        // eslint-disable-next-line @typescript-eslint/no-misused-promises, @typescript-eslint/promise-function-async
                        <MenuItem key={'previewInThisChannelKey'} icon={<OpenRegular />} onClick={() => checkPreviewMessage(item.id)}>
                          {t('PreviewInThisChannel')}
                        </MenuItem>
                      }
                      <MenuItem
                        icon={<EditRegular />}
                        key={'editMessageKey'}
                        onClick={() => {
                          onOpenTaskModule(editUrl(item.id), t('EditMessage'));
                        }}
                      >
                        {t('Edit')}
                      </MenuItem>
                      {
                        // eslint-disable-next-line @typescript-eslint/no-misused-promises, @typescript-eslint/promise-function-async
                        <MenuItem key={'duplicateKey'} icon={<DocumentCopyRegular />} onClick={() => duplicateDraftMessage(item.id)}>
                          {t('Duplicate')}
                        </MenuItem>
                      }
                      {
                        // eslint-disable-next-line @typescript-eslint/no-misused-promises, @typescript-eslint/promise-function-async
                        <MenuItem key={'deleteKey'} icon={<DeleteRegular />} onClick={() => deleteDraftMessage(item.id)}>
                          {t('Delete')}
                        </MenuItem>
                      }
                    </MenuList>
                  </MenuPopover>
                </Menu>
              </TableCellLayout>
            </TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
};
