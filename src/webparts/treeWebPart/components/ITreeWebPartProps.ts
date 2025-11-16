export interface ITreeWebPartProps {
  description: string;
  listName: string;
  rootNodeValue: string;
  childColumns: string[];
  nodeColors: Record<number, string>; // Add this line for nodeColors
  tooltipFields: Record<number, string[]>; // Correct type
  columnDisplayNames: { [key: string]: string };
  environmentMessage: string;
  userDisplayName: string;
  columnsMetadata: { [key: string]: { isLookup?: boolean; lookupField?: string; indexed?: boolean; relatedField?: string } };
  fetchLookupValue: (listId: string, itemId: number, fieldName: string) => Promise<string>;
  filterColumn?: string; // Filter column property
  filterText?: string; // Filter text property
  filterOperator?: 'AND' | 'OR'; // Filter operator for multiple values
}