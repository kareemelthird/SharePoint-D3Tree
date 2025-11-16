import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneDropdown, PropertyPaneTextField, IPropertyPaneDropdownOption, PropertyPaneChoiceGroup } from '@microsoft/sp-property-pane';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import TreeWebPart from './components/TreeWebPart';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { IFieldInfo } from '@pnp/sp/fields/types';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
  
interface IExtendedFieldInfo extends IFieldInfo {
  LookupList?: string;
  SchemaXml: string;
}


export interface ITreeWebPartWebPartProps {
  description: string;
  listName: string;
  rootNodeValue: string;
  nodeColors: Record<number, string>;
  childColumns: string[];
  tooltipFields: Record<number, string[]>; // Level-based tooltips
  filterColumn?: string; // Filter column property
  filterText?: string; // Custom filter text property
  filterOperator?: 'AND' | 'OR'; // Filter operator for multiple values
}

export default class TreeWebPartWebPart extends BaseClientSideWebPart<ITreeWebPartWebPartProps> {
  private _listNames: string[] = [];
  private _isListNamesLoaded: boolean = false;
  private _columns: IPropertyPaneDropdownOption[] = [];
  private _columnDisplayNames: { [key: string]: string } = {};
  private _columnsMetadata: { [key: string]: {
    isLookup: boolean;
    lookupList?: string;
    lookupField?: string;
    relatedField?: string;
    indexed?: boolean;
  } } = {};

  public async onInit(): Promise<void> {
    await super.onInit();
    try {
      sp.setup({
        sp: {
          baseUrl: this.context.pageContext.web.absoluteUrl,
        },
      });
      const lists = await sp.web.lists.select('Title', 'BaseTemplate').get();
      this._listNames = lists
        .filter(list => list.BaseTemplate === 100)
        .map((list) => list.Title);
      this._isListNamesLoaded = true;
      if (this.properties.listName) {
        await this._fetchListColumns(this.properties.listName);
      }
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Error fetching list names:', error);
    }
  }

  private async _fetchListColumns(listName: string): Promise<void> {
    try {
      const list = sp.web.lists.getByTitle(listName);
      const listFields: IExtendedFieldInfo[] = await list.fields.select(
        'Title', 'InternalName', 'Hidden', 'TypeAsString', 'LookupList', 'SchemaXml'
      ).get();

      this._columns = [];
      this._columnsMetadata = {};

      console.log('Fetched list fields:', listFields.length);

      listFields.forEach(field => {
        if (field.Hidden) return;

        console.log(`Processing field: ${field.InternalName} (${field.TypeAsString})`);

        if (field.TypeAsString === 'Lookup') {
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(field.SchemaXml, "text/xml");
          
          const showField = xmlDoc.documentElement.getAttribute('ShowField') || 'Title';
          const relatedField = xmlDoc.documentElement.getAttribute('RelatedField');
          const indexed = xmlDoc.documentElement.getAttribute('Indexed') === 'TRUE';

          console.log(`  Lookup field details: showField=${showField}, relatedField=${relatedField}, indexed=${indexed}`);

          // Check if this is a projected field (contains _x003a_ which is encoded colon)
          const isProjectedField = field.InternalName.includes('_x003a_');

          if (isProjectedField) {
            // This is a projected field created by SharePoint for indexed lookups
            // These should be treated as direct selectable fields, not expandable lookups
            this._columnsMetadata[field.InternalName] = {
              isLookup: false, // Treat as regular field for REST API purposes
              lookupList: field.LookupList,
              lookupField: showField,
              relatedField: relatedField || undefined,
              indexed: true
            };

            this._columns.push({
              key: field.InternalName,
              text: field.Title
            });
            this._columnDisplayNames[field.InternalName] = field.Title;

            console.log(`  Added projected field as regular field: ${field.InternalName}`);

          } else if (indexed) {
            // This is the main indexed lookup field
            this._columnsMetadata[field.InternalName] = {
              isLookup: true,
              lookupList: field.LookupList,
              lookupField: showField,
              relatedField: relatedField || undefined,
              indexed: true
            };

            this._columns.push({
              key: field.InternalName,
              text: field.Title
            });
            this._columnDisplayNames[field.InternalName] = field.Title;

            console.log(`  Added main indexed lookup: ${field.InternalName}`);

          } else {
            // Non-indexed lookup handling
            this._columnsMetadata[field.InternalName] = {
              isLookup: true,
              lookupList: field.LookupList,
              lookupField: showField,
              relatedField: relatedField || undefined,
              indexed: false
            };

            this._columns.push({
              key: field.InternalName,
              text: field.Title
            });
            this._columnDisplayNames[field.InternalName] = field.Title;

            console.log(`  Added non-indexed lookup: ${field.InternalName}`);
          }
        } else {
          this._columnsMetadata[field.InternalName] = { isLookup: false };
          
          this._columns.push({
            key: field.InternalName,
            text: field.Title
          });
          this._columnDisplayNames[field.InternalName] = field.Title;
        }
      });

      console.log('Final columns metadata:', this._columnsMetadata);
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Error fetching list columns:', error);
    }
  }

  private async _fetchLookupValue(listId: string, itemId: number, fieldName: string): Promise<string> {
    try {
      const fieldMetadata = this._columnsMetadata[fieldName];

      if (fieldMetadata.isLookup) {
        // Handle indexed lookups differently
        if (fieldMetadata.indexed) {
          // Get the ID from the related field (e.g., DepartmentId)
          const idField = fieldMetadata.relatedField;
          if (!idField) return '';
          const item = await sp.web.lists.getById(listId).items.getById(itemId)
            .select(idField)
            .get();

          const lookupItemId = item[idField];

          if (lookupItemId && fieldMetadata.lookupList && fieldMetadata.lookupField) {
            const lookupItem = await sp.web.lists.getById(fieldMetadata.lookupList).items.getById(lookupItemId)
              .select(fieldMetadata.lookupField)
              .get();
            return lookupItem[fieldMetadata.lookupField] || '';
          }
          return '';
        } else {
          // Original handling for non-indexed lookups
          const item = await sp.web.lists.getById(listId).items.getById(itemId)
            .select(`${fieldName}/Title`, `${fieldName}Id`)
            .expand(fieldName)
            .get();

          const lookupValue = item[fieldName]?.Title || '';

          if (fieldMetadata.lookupList && fieldMetadata.lookupField) {
            const lookupItemId = item[`${fieldName}Id`];
            if (lookupItemId) {
              const lookupItem = await sp.web.lists.getById(fieldMetadata.lookupList).items.getById(lookupItemId)
                .select(fieldMetadata.lookupField)
                .get();
              return `${lookupValue} - ${lookupItem[fieldMetadata.lookupField]}`;
            }
          }
          return lookupValue;
        }
      }

      // Handle non-lookup fields
      const item = await sp.web.lists.getById(listId).items.getById(itemId)
        .select(fieldName)
        .get();
      return item[fieldName] || '';
    } catch (error) {
      console.error(`Error fetching lookup value for field '${fieldName}' on item ID ${itemId}:`, error);
      return '';
    }
  }

  private _getColumnsOptions(): IPropertyPaneDropdownOption[] {
    return this._columns;
  }

  public async render(): Promise<void> {
    const element: React.ReactElement = React.createElement(TreeWebPart, {
      description: this.properties.description,
      listName: this.properties.listName,
      rootNodeValue: this.properties.rootNodeValue,
      nodeColors: this.properties.nodeColors || {},
      childColumns: this.properties.childColumns || [],
      tooltipFields: this.properties.tooltipFields || {},
      columnsMetadata: this._columnsMetadata,
      columnDisplayNames: this._columnDisplayNames,
      environmentMessage: 'Your environment message here',
      userDisplayName: this.context.pageContext.user.displayName,
      fetchLookupValue: this._fetchLookupValue.bind(this),
      filterColumn: this.properties.filterColumn, // Pass filterColumn to React component
      filterText: this.properties.filterText, // Pass filterText to React component
      filterOperator: this.properties.filterOperator // Pass filterOperator to React component
    });
    ReactDOM.render(element, this.domElement);
  }

  public onDispose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): Promise<void> {
    if (propertyPath === 'listName' && newValue !== oldValue) {
      await this._fetchListColumns(newValue as string);
    }
    if (propertyPath.startsWith('nodeColors')) {
      const level = parseInt(propertyPath.split('.')[1], 10);
      if (!isNaN(level)) {
        this.properties.nodeColors[level] = newValue as string;
      }
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): import('@microsoft/sp-property-pane').IPropertyPaneConfiguration {
    const columnOptions = this._getColumnsOptions();
    const maxLevels = 7; // Match your child column levels
    // Add color picker groups
    const colorGroups = Array.from({ length: maxLevels }, (_, i) => {
      const level = i + 1;
      return {
        groupName: `Level ${level} Color`,
        groupFields: [
          PropertyFieldColorPicker(`nodeColors.${level}`, {
            label: `Level ${level} Node Color`,
            selectedColor: this.properties.nodeColors?.[level] || '#0078d4',
            onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
            properties: this.properties,
            disabled: false,
            debounce: 300, // Delay before applying changes
            alphaSliderHidden: true, // Hide alpha slider if not needed
            style: PropertyFieldColorPickerStyle.Full, // Full or Inline display
            iconName: 'Precipitation', // Optional icon for inline mode
            key: `colorPicker-${level}`
          })
        ]
      };
    });
    // Create tooltip groups for each level
    const tooltipGroups = Array.from({ length: maxLevels }, (_, i) => {
      const level = i + 1;
      return {
        groupName: `Level ${level} Tooltips`,
        groupFields: [
          PropertyPaneDropdown(`tooltipFields.${level}.0`, {
            label: `Level ${level} - First Tooltip Field`,
            options: columnOptions,
            selectedKey: this.properties.tooltipFields?.[level]?.[0]
          }),
          PropertyPaneDropdown(`tooltipFields.${level}.1`, {
            label: `Level ${level} - Second Tooltip Field`,
            options: columnOptions,
            selectedKey: this.properties.tooltipFields?.[level]?.[1]
          }),
          PropertyPaneDropdown(`tooltipFields.${level}.2`, {
            label: `Level ${level} - Third Tooltip Field`,
            options: columnOptions,
            selectedKey: this.properties.tooltipFields?.[level]?.[2]
          })
        ]
      };
    });
    return {
      pages: [
        {
          header: { description: 'D3K3 Tree Web Part'},
          groups: [
            {
              groupName: 'List Configuration',
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: 'Select a SharePoint List',
                  options: this._isListNamesLoaded
                    ? this._listNames.map((listName: string) => ({ key: listName, text: listName }))
                    : [],
                  selectedKey: this.properties.listName,
                }),
                PropertyPaneDropdown('filterColumn', {
                  label: 'Select Filter Column (Optional)',
                  options: columnOptions,
                  selectedKey: this.properties.filterColumn
                }),
                PropertyPaneTextField('filterText', {
                  label: 'Enter Filter Text (Optional)',
                  value: this.properties.filterText || '',
                  description: 'Enter text to filter the selected column. Separate multiple values with commas for OR filtering (e.g., "New,Modified").',
                  placeholder: 'e.g., New,Modified or Active,Pending'
                }),
                PropertyPaneChoiceGroup('filterOperator', {
                  label: 'Filter Operator (for multiple values)',
                  options: [
                    { key: 'OR', text: 'OR (match any value)' },
                    { key: 'AND', text: 'AND (match all values)' }
                  ]
                }),
                PropertyPaneTextField('rootNodeValue', {
                  label: 'Enter Root Node Value',
                  value: this.properties.rootNodeValue,
                }),
                // Child column levels...
                ...Array.from({ length: maxLevels }, (_, i) => 
                  PropertyPaneDropdown(`childColumns[${i}]`, {
                    label: `Select Child Column Level ${i + 1}`,
                    options: columnOptions,
                    selectedKey: this.properties.childColumns?.[i]
                  })
                )
              ],
            },
            // Add tooltip groups
            ...tooltipGroups,
            ...colorGroups
          ],
        },
      ],
    };
  }
}