import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';

interface IMultiSelectProps {
  label: string;
  options: { key: string; text: string }[];
  selectedKeys: string[];
  onChange: (selectedKeys: string[]) => void;
}

const MultiSelect: React.FC<IMultiSelectProps> = ({ label, options, selectedKeys = [], onChange }): JSX.Element => {
  const onCheckboxChange = (key: string, checked: boolean | undefined): void => {
    const newSelectedKeys = checked ? [...selectedKeys, key] : selectedKeys.filter(k => k !== key);
    onChange(newSelectedKeys);
  };

  return (
    <div>
      <label>{label}</label>
      {options.map(option => (
        <Checkbox
          key={option.key}
          label={option.text}
          checked={selectedKeys.includes(option.key)}
          onChange={(e, checked) => onCheckboxChange(option.key, checked)}
        />
      ))}
    </div>
  );
};

export default MultiSelect;