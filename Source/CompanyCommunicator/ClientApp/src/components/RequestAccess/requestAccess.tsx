// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./requestAccess.scss";
import React from "react";
import { useTranslation } from "react-i18next";
import { RouteComponentProps } from "react-router-dom";
import { Button, Text } from "@fluentui/react-components";
import * as microsoftTeams from "@microsoft/teams-js";
import i18n from "../../i18n";

const RequestAccess: React.FunctionComponent<RouteComponentProps> = (props) => {
  const { t } = useTranslation();
  const errorMessage = t("ForbiddenErrorMessage");

  function onRequestAccess() {
    console.log("Implement logic here");
  }

  return (
    <div className="sign-in-content-container">
      <Text className="info-text" size={500}>
        {errorMessage}
      </Text>
      <div className="space"></div>
      <Button appearance="primary" className="sign-in-button" onClick={onRequestAccess}>
        {t("RequestAccess")}
      </Button>
    </div>
  );
};

export default RequestAccess;
