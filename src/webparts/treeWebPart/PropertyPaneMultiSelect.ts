import * as React from 'react';
import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import { IPropertyPaneMultiSelectProps, IPropertyPaneMultiSelectInternalProps } from './IPropertyPaneMultiSelectProps';
import MultiSelect from './components/MultiSelect';
import * as ReactDOM from 'react-dom';

export class PropertyPaneMultiSelect implements IPropertyPaneField<IPropertyPaneMultiSelectProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneMultiSelectInternalProps;

  constructor(targetProperty: string, properties: IPropertyPaneMultiSelectProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.key,
      label: properties.label,
      options: properties.options,
      selectedKeys: properties.selectedKeys || [],
      onPropertyChange: properties.onPropertyChange,
      onRender: this.onRender.bind(this)
    };
  }

  private onRender(elem: HTMLElement): void {
    const element = React.createElement(MultiSelect, {
      label: this.properties.label,
      options: this.properties.options,
      selectedKeys: this.properties.selectedKeys,
      onChange: this.onChange.bind(this)
    });
    ReactDOM.render(element, elem);
  }

  public dispose(): void {
    // Unmount the React component to avoid memory leaks
    const elem = document.querySelector(`[data-automation-id='${this.targetProperty}']`);
    if (elem) {
      ReactDOM.unmountComponentAtNode(elem as HTMLElement);
    }
  }

  private onChange(selectedKeys: string[]): void {
    this.properties.selectedKeys = selectedKeys;
    this.properties.onPropertyChange(this.targetProperty, selectedKeys);
  }
}