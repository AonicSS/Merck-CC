// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { useTranslation } from 'react-i18next';
import { Button, Text } from '@fluentui/react-components';

const RequestAccess = () => {
  const { t } = useTranslation();
  const errorMessage = t('ForbiddenErrorMessage');

  function onRequestAccess() {
    console.log('Implement logic here');
  }

  return (
    <div className='sign-in-content-container'>
      <Text className='info-text' size={500}>
        {errorMessage}
      </Text>
      <div className='space'></div>
      <Button appearance='primary' className='sign-in-button' onClick={onRequestAccess}>
        {t('RequestAccess')}
      </Button>
    </div>
  );
};

export default RequestAccess;
