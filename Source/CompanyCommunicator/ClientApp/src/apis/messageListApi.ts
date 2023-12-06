// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { getBaseUrl } from '../configVariables';
import { IDeleteMessageRequest } from '../models/deleteMessages';
import apiCall from './apiDecorator';

const baseAxiosUrl = getBaseUrl() + '/api';

export const getSentNotifications = async (): Promise<any> => {
  const url = baseAxiosUrl + '/sentnotifications';
  return await apiCall.getJsonResponse(url);
};

export const getDraftNotifications = async (): Promise<any> => {
  const url = baseAxiosUrl + '/draftnotifications';
  return await apiCall.getJsonResponse(url);
};

export const verifyGroupAccess = async (): Promise<any> => {
  const url = baseAxiosUrl + '/groupdata/verifyaccess';
  return await apiCall.getTextResponse(url);
};

export const getGroups = async (id: number): Promise<any> => {
  const url = `${baseAxiosUrl}/groupdata/${id}`;
  return await apiCall.getJsonResponse(url);
};

export const searchGroups = async (query: string): Promise<any> => {
  const url = `${baseAxiosUrl}/groupdata/search/${query}`;
  return await apiCall.getTextResponse(url);
};

export const getTeams = async (): Promise<any> => {
  const url = baseAxiosUrl + '/teamdata';
  return await apiCall.getJsonResponse(url);
};

export const getDraftNotification = async (id: number): Promise<any> => {
  const url = `${baseAxiosUrl}/draftnotifications/${id}`;
  return await apiCall.getJsonResponse(url);
};

export const exportNotification = async (payload: any): Promise<any> => {
  const url = baseAxiosUrl + '/exportnotification/export';
  return await apiCall.putAndGetTextResponse(url, payload);
};

export const getSentNotification = async (id: number): Promise<any> => {
  const url = `${baseAxiosUrl}/sentnotifications/${id}`;
  return await apiCall.getJsonResponse(url);
};

export const deleteDraftNotification = async (id: number): Promise<any> => {
  const url = `${baseAxiosUrl}/draftnotifications/${id}`;
  return await apiCall.deleteAndGetTextResponse(url);
};

export const duplicateDraftNotification = async (id: number): Promise<any> => {
  const url = `${baseAxiosUrl}/draftnotifications/duplicates/${id}`;
  return await apiCall.postAndGetTextResponse(url);
};

export const sendDraftNotification = async (payload: any): Promise<any> => {
  const url = baseAxiosUrl + '/sentnotifications';
  return await apiCall.postAndGetTextResponse(url, payload);
};

export const updateDraftNotification = async (payload: any): Promise<any> => {
  const url = baseAxiosUrl + '/draftnotifications';
  return await apiCall.putAndGetJsonResponse(url, payload);
};

export const createDraftNotification = async (payload: any): Promise<any> => {
  const url = baseAxiosUrl + '/draftnotifications';
  return await apiCall.postAndGetJsonResponse(url, payload);
};

export const cancelSentNotification = async (id: number): Promise<any> => {
  const url = `${baseAxiosUrl}/sentnotifications/cancel/${id}`;
  return await apiCall.postAndGetTextResponse(url);
};

export const getConsentSummaries = async (id: number): Promise<any> => {
  const url = `${baseAxiosUrl}/draftnotifications/consentSummaries/${id}`;
  return await apiCall.getJsonResponse(url);
};

export const sendPreview = async (payload: any): Promise<any> => {
  const url = baseAxiosUrl + '/draftnotifications/previews';
  return await apiCall.postAndGetTextResponse(url, payload);
};

export const getDeletedMessages = async (): Promise<any> => {
  const url = baseAxiosUrl + '/deletemessages';
  return await apiCall.getJsonResponse(url);
};

export const deleteMessages = async (payload: IDeleteMessageRequest): Promise<any> => {
  const url = baseAxiosUrl + '/deletemessages';
  return await apiCall.postAndGetJsonResponse(url, payload);
};

export const getAuthenticationConsentMetadata = async (windowLocationOriginDomain: string, loginHint: string): Promise<any> => {
  const url = `${baseAxiosUrl}/authenticationMetadata/consentUrl?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${loginHint}`;
  return await apiCall.getJsonResponse(url);
};

export const getScheduledDraftNotifications = async (): Promise<any> => {
  const url = baseAxiosUrl + '/draftnotifications/scheduledDrafts';
  return await apiCall.getJsonResponse(url);
};

// custom changes

// get unit notification data

export const getUnitDraftNotification = async (id: string): Promise<any> => {
  const url = baseAxiosUrl + "/notifications/draft/" + id;
  return await apiCall.getJsonResponse(url);
}

export const getUnitSentNotification = async (id: string): Promise<any> => {
  const url = baseAxiosUrl + "/notifications/sent/" + id;
  return await apiCall.getJsonResponse(url);
}

// get all units
export const getUnits = async (): Promise<any> => {
  const url = baseAxiosUrl + "/unitData";
  return await apiCall.getJsonResponse(url);
};

// get user units
export const getUserUnits = async (userId: string): Promise<any> => {
  const url = baseAxiosUrl + `/unitData/user/${userId}`;
  return await apiCall.getJsonResponse(url);
}

// get a single unit
export const getUnit = async (id: string): Promise<any> => {
  const url = baseAxiosUrl + "/unitData/" + id;
  if (id === "new") {
    return [];
  }
  return await apiCall.getJsonResponse(url);
};

// update a single unit, return updated data
export const updateUnit = async (unitData: { id: string; name: string; users: any; groups: any }): Promise<any> => {
  const url = baseAxiosUrl + "/unitData";
  const UserIds = unitData.users.map((user: any) => user.id);
  const GroupIds = unitData.groups.map((group: any) => group.id);

  const postData = {
    id: unitData.id,
    name: unitData.name,
    UserIds,
    GroupIds,
  };
  if (postData.id) {
    return await apiCall.putAndGetJsonResponse(url, postData);
  } else {
    return await apiCall.postAndGetJsonResponse(url, postData);
  }
};


// delete a single unit, return 200 ok
export const deleteUnit = async (id: string): Promise<any> => {
  const url = baseAxiosUrl + "/unitData/" + id;
  const response = await apiCall.deleteAndGetJsonResponse(url);
  return response;
};


// get all users
export const getUsers = async (query: string): Promise<any> => {
  const url = baseAxiosUrl + "/userData/search/startswith(userPrincipalName,'" + query + "')";
  return await apiCall.getJsonResponse(url);
};


// get single user
export const getUser = async (mail: string): Promise<any> => {
  const url = baseAxiosUrl + "/userData/" + mail;
  return await apiCall.getJsonResponse(url);
};