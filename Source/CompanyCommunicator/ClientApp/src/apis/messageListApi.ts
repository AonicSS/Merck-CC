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

const mockedMemberData = [
  { id: 1, email: "hans@wurst@aonicdemotenant.com", name: "Hans Wurst" },
  { id: 2, email: "user2@example.com", name: "John Doe" },
  { id: 3, email: "user3@example.com", name: "Jane Doe" }
];

const mockedGroupData = [
  { id: 1, name: "Test_Grp_1" },
  { id: 2, name: "Test_Grp_2" },
  { id: 3, name: "Test_Grp_3" }
];

const mockedUnitData = {
  "data": [
    { id: 1, name: "Test_Unit_1", members: mockedMemberData, groups: mockedGroupData },
    { id: 2, name: "Test_Unit_2", members: mockedMemberData, groups: mockedGroupData },
    { id: 3, name: "Test_Unit_3", members: mockedMemberData, groups: mockedGroupData },
    { id: 4, name: "Test_Unit_4", members: mockedMemberData, groups: mockedGroupData }
  ]
};


// get all units
export const getUnits = async (): Promise<any> => {
  let url = baseAxiosUrl + "/units";

  return mockedUnitData;
};

// get a single unit
export const getUnit = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  const unitData = [...mockedUnitData.data].find(item => item.id === id);
  return unitData;
};


// update a single unit, return updated unit
export const updateUnit = async (id: number, name: string): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;

  const unitToUpdate = mockedUnitData.data.find(unit => unit.id === id);
  if (unitToUpdate) {
    const updatedUnit = { ...unitToUpdate, name };
    mockedUnitData.data = mockedUnitData.data.map(unit =>
      unit.id === id ? updatedUnit : unit
    );
    console.log(mockedUnitData);
    return mockedUnitData;
  } else {
    throw new Error(`Unit with ID ${id} not found.`);
  }
};



export const addUnitMember = async (id: number, newMember: { id: number, name: string, email: string }): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;
  const unitToUpdate = mockedUnitData.data.find(unit => unit.id === id);

  if (unitToUpdate) {
    const updatedUnit = { ...unitToUpdate };
    updatedUnit.members = [...updatedUnit.members, newMember];
    mockedUnitData.data = mockedUnitData.data.map(unit =>
      unit.id === id ? updatedUnit : unit
    );
    return mockedUnitData; // Return the updated data
  } else {
    throw new Error(`Unit with ID ${id} not found.`);
  }
};

export const deleteUnitMember = async (id: number, userId: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;
  const unitToUpdate = mockedUnitData.data.find(unit => unit.id === id);

  if (unitToUpdate) {
    const updatedUnit = { ...unitToUpdate };
    updatedUnit.members = updatedUnit.members.filter(member => member.id !== userId);
    mockedUnitData.data = mockedUnitData.data.map(unit =>
      unit.id === id ? updatedUnit : unit
    );
    return mockedUnitData; // Return the updated data
  } else {
    throw new Error(`Unit with ID ${id} not found.`);
  }

};

export const addUnitGroup = async (id: number, newGroup: { id: number, name: string }): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;
  const unitToUpdate = mockedUnitData.data.find(unit => unit.id === id);

  if (unitToUpdate) {
    const updatedUnit = { ...unitToUpdate };
    updatedUnit.groups = [...updatedUnit.groups, newGroup];
    mockedUnitData.data = mockedUnitData.data.map(unit =>
      unit.id === id ? updatedUnit : unit
    );
    return mockedUnitData; // Return the updated data
  } else {
    throw new Error(`Unit with ID ${id} not found.`);
  }

};

export const deleteUnitGroup = async (id: number, groupId: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;
  const unitToUpdate = mockedUnitData.data.find(unit => unit.id === id);

  if (unitToUpdate) {
    const updatedUnit = { ...unitToUpdate };
    updatedUnit.groups = updatedUnit.groups.filter(group => group.id !== groupId);
    mockedUnitData.data = mockedUnitData.data.map(unit =>
      unit.id === id ? updatedUnit : unit
    );
    return mockedUnitData; // Return the updated data
  } else {
    throw new Error(`Unit with ID ${id} not found.`);
  }
};


// delete a single unit, return 200 ok
export const deleteUnit = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/units" + id;
  const updatedData = mockedUnitData.data.filter(unit => unit.id !== id);

  if (updatedData.length !== mockedUnitData.data.length) {
    mockedUnitData.data = updatedData;
    console.log(mockedUnitData);
    return mockedUnitData;
  } else {
    throw new Error(`Unit with ID ${id} not found.`);
  }
};


const mockedAllMemberData = {
  "data": [
    { id: 1, email: "hans@wurst@aonicdemotenant.com", name: "Hans Wurst" },
    { id: 2, email: "user2@example.com", name: "John Doe" },
    { id: 3, email: "user3@example.com", name: "Jane Doe" },
    { id: 4, email: "user3@example.com", name: "Jack Doe" },
    { id: 5, email: "user3@example.com", name: "Jill Doe" },
  ]
};

const mockedAllGroupData = {
  "data": [
    { id: 1, name: "Test_Grp_1" },
    { id: 2, name: "Test_Grp_2" },
    { id: 3, name: "Test_Grp_3" },
    { id: 4, name: "Test_Grp_4" },
    { id: 5, name: "Test_Grp_5" },
    { id: 6, name: "Test_Grp_6" }
  ]
};

// get all AD Groups,
export const getADGroups = async (): Promise<any> => {
  let url = baseAxiosUrl + "/adGroups";

  return mockedAllGroupData;
};

// get all users
export const getUsers = async (): Promise<any> => {
  let url = baseAxiosUrl + "/users";

  return mockedAllMemberData;
};