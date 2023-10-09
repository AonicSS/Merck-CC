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

const mockedMemberData = {
  "data": [
    { id: 1, email: "hans@wurst@aonicdemotenant.com", name: "Hans Wurst" },
    { id: 2, email: "user2@example.com", name: "John Doe" },
    { id: 3, email: "user3@example.com", name: "Jane Doe" }
  ]
};

const mockedGroupData = {
  "data": [
    { id: 1, name: "Test_Grp_1" },
    { id: 2, name: "Test_Grp_2" },
    { id: 3, name: "Test_Grp_3" }
  ]
};

const mockedUnitData = {
  "data": [
    { id: 1, name: "Test_Unit_1", members: mockedMemberData, groups: mockedGroupData },
    { id: 2, name: "Test_Unit_2", members: mockedMemberData, groups: mockedGroupData },
    { id: 3, name: "Test_Unit_3", members: mockedMemberData, groups: mockedGroupData },
    { id: 4, name: "Test_Unit_4", members: mockedMemberData, groups: mockedGroupData }
  ]
};


export const getUnits = async (): Promise<any> => {
  let url = baseAxiosUrl + "/units";

  return mockedUnitData;
};

export const getUnit = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  return mockedUnitData.data[0];
};

export const updateUnit = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  return mockedUnitData;
};


export const addUnitMember = async (id: number, newMember: { id: number, name: string, email: string }): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  const updatedData = [...mockedMemberData.data];
  updatedData.push(newMember);
  mockedMemberData.data = updatedData;
  return mockedMemberData;


};

export const deleteUnitMember = async (id: number, userId: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  const randomIndex = Math.floor(Math.random() * mockedMemberData.data.length);
  const updatedData = [...mockedMemberData.data];
  updatedData.splice(randomIndex, 1);
  mockedMemberData.data = updatedData;
  return mockedMemberData;


};

export const addUnitGroup = async (id: number, newGroup: { id: number, name: string }): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  const updatedData = [...mockedGroupData.data];
  updatedData.push(newGroup);
  mockedGroupData.data = updatedData;
  return mockedGroupData;


};

export const deleteUnitGroup = async (id: number, groupId: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  const randomIndex = Math.floor(Math.random() * mockedGroupData.data.length);
  const updatedData = [...mockedGroupData.data];
  updatedData.splice(randomIndex, 1);
  mockedGroupData.data = updatedData;
  return mockedGroupData;


};