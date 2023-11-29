// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { getBaseUrl } from "../configVariables";
import axios from "./axiosJWTDecorator";

let baseAxiosUrl = getBaseUrl() + "/api";

export const getSentNotifications = async (): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications";
  return await axios.get(url);
};

export const getDraftNotifications = async (): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications";
  return await axios.get(url);
};

export const verifyGroupAccess = async (): Promise<any> => {
  let url = baseAxiosUrl + "/groupdata/verifyaccess";
  return await axios.get(url, false);
};

export const getGroups = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/groupdata/" + id;
  return await axios.get(url);
};

export const searchGroups = async (query: string): Promise<any> => {
  let url = baseAxiosUrl + "/groupdata/search/" + query;
  return await axios.get(url);
};

export const exportNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/exportnotification/export";
  return await axios.put(url, payload);
};

export const getSentNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications/" + id;
  return await axios.get(url);
};

export const getDraftNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/" + id;
  return await axios.get(url);
};

export const getUnitDraftNotification = async (id: string): Promise<any> => {
  let url = baseAxiosUrl + "/notifications/draft/" + id;
  return await axios.get(url);
}

export const getUnitSentNotification = async (id: string): Promise<any> => {
  let url = baseAxiosUrl + "/notifications/sent/" + id;
  return await axios.get(url);
}

export const deleteDraftNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/" + id;
  return await axios.delete(url);
};

export const duplicateDraftNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/duplicates/" + id;
  return await axios.post(url);
};

export const sendDraftNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications";
  return await axios.post(url, payload);
};

export const updateDraftNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications";
  return await axios.put(url, payload);
};

export const createDraftNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications";
  return await axios.post(url, payload);
};

export const getTeams = async (): Promise<any> => {
  let url = baseAxiosUrl + "/teamdata";
  return await axios.get(url);
};

export const cancelSentNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications/cancel/" + id;
  return await axios.post(url);
};

export const getConsentSummaries = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/consentSummaries/" + id;
  return await axios.get(url);
};

export const sendPreview = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/previews";
  return await axios.post(url, payload);
};

export const getAuthenticationConsentMetadata = async (
  windowLocationOriginDomain: string,
  login_hint: string
): Promise<any> => {
  let url = `${baseAxiosUrl}/authenticationMetadata/consentUrl?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${login_hint}`;
  return await axios.get(url, undefined, false);
};


// New api endpoints for CC

// get all units
export const getUnits = async (): Promise<any> => {
  let url = baseAxiosUrl + "/unitData";
  return await axios.get(url);
};

// get user units
export const getUserUnits = async (userId:string): Promise<any> => {
  let url = baseAxiosUrl + `/unitData/user/${userId}`;
  return await axios.get(url);
}

// get a single unit
export const getUnit = async (id: string): Promise<any> => {
  let url = baseAxiosUrl + "/unitData/" + id;
  if (id === "new") {
    return [];
  }
  return await axios.get(url);
};


// update a single unit, return updated data
export const updateUnit = async (unitData: { id: string; name: string; users: any; groups: any }): Promise<any> => {
  let url = baseAxiosUrl + "/unitData";
  const UserIds = unitData.users.map((user: any) => user.id);
  const GroupIds = unitData.groups.map((group: any) => group.id);

  const postData = {
    id: unitData.id,
    name: unitData.name,
    UserIds,
    GroupIds,
  };
  console.log(postData);

  if (postData.id) {
    return await axios.put(url, postData);
  } else {
    return await axios.post(url, postData);
  }
};


// delete a single unit, return 200 ok
export const deleteUnit = async (id: string): Promise<any> => {
  let url = baseAxiosUrl + "/unitData/" + id;
  return await axios.delete(url);
};


// get all users
export const getUsers = async (query: string): Promise<any> => {
  let url = baseAxiosUrl + "/userData/search/startswith(userPrincipalName,'" + query + "')";
  return await axios.get(url);
};

export const getUser = async (mail: string): Promise<any> => {
  let url = baseAxiosUrl + "/userData/" + mail;
  return await axios.get(url);
};