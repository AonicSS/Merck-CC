// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createSlice } from "@reduxjs/toolkit";

export interface MessagesState {
  draftMessages: { action: string; payload: [] };
  sentMessages: { action: string; payload: [] };
  selectedMessage: { action: string; payload: {} };
  teamsData: { action: string; payload: any[] };
  groups: { action: string; payload: any[] };
  queryGroups: { action: string; payload: any[] };
  verifyGroup: { action: string; payload: boolean };
  isDraftMessagesFetchOn: { action: string; payload: boolean };
  isSentMessagesFetchOn: { action: string; payload: boolean };
  units: { action: string, payload: any[] };
  unit: { action: string, payload: {} };
  unitMembers: { action: string; payload: any[] };
  unitGroups: { action: string; payload: any[] };
  users: { action: string; payload: any[] };
  adGroups: { action: string; payload: any[] };
}

const initialState: MessagesState = {
  draftMessages: { action: "FETCH_DRAFT_MESSAGES", payload: [] },
  sentMessages: { action: "FETCH_MESSAGES", payload: [] },
  selectedMessage: { action: "MESSAGE_SELECTED", payload: [] },
  teamsData: { action: "GET_TEAMS_DATA", payload: [] },
  groups: { action: "GET_GROUPS", payload: [] },
  queryGroups: { action: "SEARCH_GROUPS", payload: [] },
  verifyGroup: { action: "VERIFY_GROUP_ACCESS", payload: false },
  isDraftMessagesFetchOn: { action: "DRAFT_MESSAGES_FETCH_STATUS", payload: false },
  isSentMessagesFetchOn: { action: "SENT_MESSAGES_FETCH_STATUS", payload: false },
  units: { action: "GET_UNITS", payload: [] },
  unit: { action: "GET_UNIT", payload: {} },
  unitMembers: { action: "GET_UNIT_MEMBERS", payload: [] },
  unitGroups: { action: "GET_UNIT_GROUPS", payload: [] },
  users: { action: "GET_UNIT_GROUPS", payload: [] },
  adGroups: { action: "GET_UNIT_GROUPS", payload: [] }
};

export const messagesSlice = createSlice({
  name: "messagesSlice",
  initialState,
  reducers: {
    draftMessages: (state, action) => {
      state.draftMessages = action.payload;
    },
    sentMessages: (state, action) => {
      state.sentMessages = action.payload;
    },
    selectedMessage: (state, action) => {
      state.selectedMessage = action.payload;
    },
    teamsData: (state, action) => {
      state.teamsData = action.payload;
    },
    groups: (state, action) => {
      state.groups = action.payload;
    },
    queryGroups: (state, action) => {
      state.queryGroups = action.payload;
    },
    verifyGroup: (state, action) => {
      state.verifyGroup = action.payload;
    },
    isDraftMessagesFetchOn: (state, action) => {
      state.isDraftMessagesFetchOn = action.payload;
    },
    isSentMessagesFetchOn: (state, action) => {
      state.isSentMessagesFetchOn = action.payload;
    },
    units: (state, action) => {
      state.units = action.payload;
    },
    unit: (state, action) => {
      state.unit = action.payload;
    },
    unitMembers: (state, action) => {
      state.unitMembers = action.payload;
    },
    unitGroups: (state, action) => {
      state.unitGroups = action.payload;
    },
    users: (state, action) => {
      state.users = action.payload;
    },
    adGroups: (state, action) => {
      state.adGroups = action.payload;
    },
  },
});

export const {
  draftMessages,
  sentMessages,
  selectedMessage,
  teamsData,
  groups,
  queryGroups,
  verifyGroup,
  isDraftMessagesFetchOn,
  isSentMessagesFetchOn,
  units,
  unit,
  unitMembers,
  unitGroups,
  users,
  adGroups
} = messagesSlice.actions;

export default messagesSlice.reducer;
