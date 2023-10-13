// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation } from "react-i18next";
import {
  Combobox,
  Option,
  useId,
} from "@fluentui/react-components";
import type { ComboboxProps } from "@fluentui/react-components";
import {
  Add24Filled
} from "@fluentui/react-icons";
 

export const ComboBox = (props: any) => {
  const options = props.options;
  const [inputText, setInputText] = React.useState<string>('');
  const [matchingOptions, setMatchingOptions] = React.useState([...options]);

  React.useEffect(() => {
    // Check if 'options' have changed before updating 'matchingOptions' to prevent re-render
    if (JSON.stringify(options) !== JSON.stringify(matchingOptions)) {
      setMatchingOptions([...options]);
    }
  }, [options, matchingOptions]);


  const comboId = useId("combo-default");
  const onChange: ComboboxProps["onChange"] = (event) => {
    const value = event.target.value.trim();
    const matches = options.filter(
      (option: any) => option.name.toLowerCase().indexOf(value.toLowerCase()) === 0
    );
    setMatchingOptions(matches);
  };
  return (
    <div style={{ display: "flex", alignItems: "center", gap: "15px", marginTop: "15px" }}>
      <Add24Filled />
      <Combobox
        aria-labelledby={comboId}
        placeholder={props.placeholder }
        onChange={onChange}
        onOptionSelect={(event, option) => {
          const selectedObject = matchingOptions.find(item => item.name === option.optionText);
          props.onSelect(selectedObject);
        }}
        input={{
          value: inputText,
          onChange: (event: React.FormEvent<HTMLInputElement>, newValue?: string) => setInputText(newValue!),
        }}
      >
        {matchingOptions.map((option) => (
          <Option key={option.id}>
            {option.name}
          </Option>
        ))}
        {matchingOptions.length === 0 ? (
          <Option key="no-results" text="">
            No results found
          </Option>
        ) : null}
      </Combobox>
    </div>
  );
};