export interface IPropertyPaneMultiSelectProps {
  key: string;
  label: string;
  options: { key: string; text: string }[];
  selectedKeys: string[];
  onPropertyChange: (propertyPath: string, newValue: string[]) => void;
}

export interface IPropertyPaneMultiSelectInternalProps extends IPropertyPaneMultiSelectProps {
  onRender: (elem: HTMLElement) => void;
}