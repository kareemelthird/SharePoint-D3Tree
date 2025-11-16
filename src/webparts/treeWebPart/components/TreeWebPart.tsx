import * as React from 'react';
import { useEffect, useState , useRef } from 'react';
import { sp } from '@pnp/sp';
import * as d3 from 'd3';
import { ITreeWebPartProps } from './ITreeWebPartProps';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faExpand, faCompress } from '@fortawesome/free-solid-svg-icons';
import styles from './TreeWebPart.module.scss';


interface D3Node {
  title: string;
  children?: D3Node[];
  level: number;
  tooltipData?: { [key: string]: string };
}

type CustomHierarchyNode = d3.HierarchyNode<D3Node> & {
  _children?: CustomHierarchyNode[];
};

const TreeWebPart: React.FC<ITreeWebPartProps> = (props) => {
  const [treeData, setTreeData] = useState<D3Node[]>([]);
  const [containerWidth, setContainerWidth] = useState<number>(0);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const updateWidth = (): void => {
      setContainerWidth(containerRef.current?.clientWidth || 800);
    };

    updateWidth();
    window.addEventListener('resize', updateWidth);

    return () => {
      window.removeEventListener('resize', updateWidth);
    };
  }, []);

  // Add filter parameter (new function to avoid breaking existing usages)
  const fetchAllItemsWithFilter = async (
    listName: string,
    childColumns: string[],
    tooltipFields: Record<number, string[]>,
    columnsMetadata: { [key: string]: { isLookup?: boolean; lookupField?: string; indexed?: boolean; relatedField?: string } },
    filter?: string
  ): Promise<Record<string, unknown>[]> => {
    const select: string[] = [];
    const expand: string[] = [];
    
    // Helper function to process field for select/expand
    const processField = (column: string): void => {
      const meta = columnsMetadata[column];
      
      // Check if this is a projected field (contains _x003a_)
      if (column.includes('_x003a_')) {
        // This is a projected field - we need to access it through the main lookup field
        // Extract the main field name (part before _x003a_)
        const mainFieldName = column.split('_x003a_')[0];
        const mainFieldMeta = columnsMetadata[mainFieldName];
        
        if (mainFieldMeta?.isLookup) {
          // Get the show field for this projected field
          const showField = meta?.lookupField || 'Title';
          select.push(`${mainFieldName}/${showField}`);
          expand.push(mainFieldName);
          
          console.log(`Added projected field through main lookup: ${mainFieldName}/${showField} (for ${column})`);
        } else {
          console.log(`Warning: Could not find main lookup field for projected field: ${column}`);
        }
        return;
      }
      
      if (meta?.isLookup) {
        if (meta.indexed) {
          // For indexed lookups, SharePoint automatically creates projected fields
          // We should request the lookup field with /Title or /[ShowField]
          const lookupField = meta.lookupField || 'Title';
          select.push(`${column}/${lookupField}`);
          expand.push(column);
          
          // Also add the ID field for debugging
          select.push(`${column}Id`);
          
          console.log(`Added indexed lookup: ${column}/${lookupField}`);
        } else {
          // Non-indexed lookup - traditional handling
          const lookupField = meta.lookupField || 'Title';
          select.push(`${column}/${lookupField}`);
          expand.push(column);
          
          console.log(`Added non-indexed lookup: ${column}/${lookupField}`);
        }
      } else {
        // Non-lookup field
        select.push(column);
        console.log(`Added regular field: ${column}`);
      }
    };

    // Add child columns to select/expand
    childColumns.forEach(processField);

    // Add tooltip fields to select/expand
    const allTooltipFields = new Set<string>();
    Object.values(tooltipFields).forEach(fields => fields.forEach(f => allTooltipFields.add(f)));
    Array.from(allTooltipFields).forEach(processField);

    // Remove duplicates
    const uniqueSelect = Array.from(new Set(select));
    const uniqueExpand = Array.from(new Set(expand));

    console.log('REST API Select fields:', uniqueSelect);
    console.log('REST API Expand fields:', uniqueExpand);

    try {
      let items: Record<string, unknown>[] = [];
      let query = sp.web.lists.getByTitle(listName).items;
      
      if (uniqueSelect.length > 0) {
        query = query.select(...uniqueSelect);
      }
      
      if (uniqueExpand.length > 0) {
        query = query.expand(...uniqueExpand);
      }
      
      query = query.top(5000);
      
      if (filter) {
        query = query.filter(filter);
      }

      console.log('Executing REST query...');
      let batch = await query.getPaged();
      items = items.concat(batch.results);
      
      while (batch.hasNext) {
        batch = await batch.getNext();
        items = items.concat(batch.results);
      }
      
      console.log(`Fetched ${items.length} items`);
      if (items.length > 0) {
        console.log('First item structure:', Object.keys(items[0]));
        console.log('Sample item data:', items[0]);
      }
      
      return items;
    } catch (error) {
      console.error('Error fetching items:', error);
      console.error('Query details - Select:', uniqueSelect, 'Expand:', uniqueExpand);
      return [];
    }
  };

  // Move buildHierarchy above useEffect to avoid 'used before defined' error
  const buildHierarchy = (
    items: Record<string, unknown>[],
    rootNodeValue: string,
    childColumns: string[],
    tooltipFields: Record<number, string[]>,
    columnsMetadata: { [key: string]: { isLookup?: boolean; lookupField?: string; indexed?: boolean; relatedField?: string } }
  ): D3Node[] => {
    const rootNodes: { [key: string]: D3Node } = {
      [rootNodeValue]: { title: rootNodeValue, children: [], level: 0 }
    };

    // Debug first item to understand field structure
    if (items.length > 0) {
      console.log('=== DEBUGGING FIRST ITEM ===');
      console.log('Available fields:', Object.keys(items[0]));
      console.log('Sample values:', items[0]);
      console.log('Child columns to process:', childColumns);
      console.log('Columns metadata:', columnsMetadata);
    }

    items.forEach((item, itemIndex) => {
      let currentNode = rootNodes[rootNodeValue];
      
      childColumns.forEach((column, levelIndex) => {
        const meta = columnsMetadata[column];
        let childValue: string = '';
        
        // Check if this is a projected field
        if (column.includes('_x003a_')) {
          // Extract the main field name and the show field
          const mainFieldName = column.split('_x003a_')[0];
          const mainFieldMeta = columnsMetadata[mainFieldName];
          const projectedFieldMeta = columnsMetadata[column];
          
          if (mainFieldMeta?.isLookup && projectedFieldMeta?.lookupField) {
            const showField = projectedFieldMeta.lookupField;
            const projectedFieldName = `${mainFieldName}/${showField}`;
            
            // Try to get the value from the expanded lookup object
            childValue = (item[projectedFieldName] as string) || 
                        (item[mainFieldName] as Record<string, string> | undefined)?.[showField] || '';
            
            // Debug logging for first few items
            if (itemIndex < 3) {
              console.log(`Item ${itemIndex}, Projected field ${column}:`);
              console.log(`  Main field: ${mainFieldName}, Show field: ${showField}`);
              console.log(`  Projected path (${projectedFieldName}):`, item[projectedFieldName]);
              console.log(`  Expanded object (${mainFieldName}):`, item[mainFieldName]);
              console.log(`  Final value:`, childValue);
            }
          }
        } else if (meta?.isLookup) {
          const lookupField = meta.lookupField || 'Title';
          const projectedFieldName = `${column}/${lookupField}`;
          
          // Try to get the value from the projected field first
          childValue = (item[projectedFieldName] as string) || '';
          
          // If that doesn't work, try the expanded object format
          if (!childValue) {
            childValue = (item[column] as Record<string, string> | undefined)?.[lookupField] || '';
          }
          
          // Debug logging for first few items
          if (itemIndex < 3) {
            console.log(`Item ${itemIndex}, Lookup field ${column}:`);
            console.log(`  Projected field (${projectedFieldName}):`, item[projectedFieldName]);
            console.log(`  Expanded object (${column}):`, item[column]);
            console.log(`  Final value:`, childValue);
          }
          
        } else {
          // Non-lookup field
          childValue = (item[column] as string) || '';
          
          if (itemIndex < 3) {
            console.log(`Item ${itemIndex}, Regular field ${column}:`, childValue);
          }
        }

        const level = levelIndex + 1;
        
        if (childValue) {
          let childNode = currentNode.children?.find((child) => child.title === childValue);
          
          if (!childNode) {
            childNode = {
              title: childValue,
              children: [],
              level: level,
              tooltipData: {}
            };

            // Process tooltip fields for this level
            const levelTooltips = tooltipFields[level] || [];
            levelTooltips.forEach(field => {
              const fieldMeta = columnsMetadata[field];
              let value = '';
              
              // Check if this is a projected field
              if (field.includes('_x003a_')) {
                // Extract the main field name and the show field
                const mainFieldName = field.split('_x003a_')[0];
                const mainFieldMeta = columnsMetadata[mainFieldName];
                const projectedFieldMeta = columnsMetadata[field];
                
                if (mainFieldMeta?.isLookup && projectedFieldMeta?.lookupField) {
                  const showField = projectedFieldMeta.lookupField;
                  const projectedFieldName = `${mainFieldName}/${showField}`;
                  
                  // Try to get the value from the expanded lookup object
                  value = (item[projectedFieldName] as string) || 
                         (item[mainFieldName] as Record<string, string> | undefined)?.[showField] || '';
                }
              } else if (fieldMeta?.isLookup) {
                const lookupField = fieldMeta.lookupField || 'Title';
                const projectedFieldName = `${field}/${lookupField}`;
                
                // Try projected field first, then expanded object
                value = (item[projectedFieldName] as string) || 
                       (item[field] as Record<string, string> | undefined)?.[lookupField] || '';
              } else {
                // Non-lookup field
                value = (item[field] as string) || '';
              }
              
              if (value) {
                childNode!.tooltipData![field] = value;
              }
            });

            if (!currentNode.children) currentNode.children = [];
            currentNode.children.push(childNode);
          }
          currentNode = childNode;
        } else {
          // Stop building hierarchy if we hit an empty value
          if (itemIndex < 3) {
            console.log(`Stopping hierarchy build for item ${itemIndex} at level ${level} - empty value for column ${column}`);
          }
          return;
        }
      });
    });

    console.log('=== HIERARCHY BUILD COMPLETE ===');
    const result = Object.values(rootNodes);
    console.log('Final hierarchy:', result);
    return result;
  };

  useEffect(() => {
    const fetchData = async (): Promise<void> => {
      try {
        if (props.listName && props.rootNodeValue && props.childColumns.length > 0) {
          // Build filter based on user input
          let filter = '';
          if (props.filterColumn && props.filterText) {
            // Split filter text by commas and trim whitespace
            const filterValues = props.filterText.split(',').map(value => value.trim()).filter(value => value);
            
            if (filterValues.length === 1) {
              // Single value filter
              filter = `${props.filterColumn} eq '${filterValues[0]}'`;
            } else if (filterValues.length > 1) {
              // Multiple values filter with OR/AND logic
              const operator = props.filterOperator || 'OR';
              const conditions = filterValues.map(value => `${props.filterColumn} eq '${value}'`);
              
              if (operator === 'OR') {
                filter = `(${conditions.join(' or ')})`;
              } else {
                filter = `(${conditions.join(' and ')})`;
              }
            }
            
            console.log('Applied filter:', filter);
          }
          const items = await fetchAllItemsWithFilter(
            props.listName,
            props.childColumns,
            props.tooltipFields,
            props.columnsMetadata as { [key: string]: { isLookup?: boolean; lookupField?: string; indexed?: boolean; relatedField?: string } },
            filter
          );
          const hierarchicalData = buildHierarchy(
            items,
            props.rootNodeValue,
            props.childColumns,
            props.tooltipFields || {},
            props.columnsMetadata as { [key: string]: { isLookup?: boolean; lookupField?: string; indexed?: boolean; relatedField?: string } }
          );
          setTreeData(hierarchicalData);
        }
      } catch (error) {
        console.error('Error fetching list items:', error);
      }
    };
    // Use fetchData().catch to handle unhandled promise lint error
    fetchData().catch(() => {});
  }, [props.listName, props.rootNodeValue, props.childColumns, props.tooltipFields, props.columnsMetadata, props.filterColumn, props.filterText, props.filterOperator]);

  // Ensure columns are loaded when listName changes
  useEffect(() => {
    if (props.listName) {
      if (typeof window !== 'undefined' && (window as unknown as { SPComponentLoader?: unknown }).SPComponentLoader) {
        // If running in SPFx, reload property pane columns
        const context = (props as unknown as { __webPartContext?: { propertyPane?: { refresh: () => void } } }).__webPartContext;
        if (context && context.propertyPane) {
          context.propertyPane.refresh();
        }
      }
      // If not, fallback: force a re-render by updating a dummy state or trigger a callback if available
    }
  }, [props.listName]);

  const d3Ref = useRef<{
    root: CustomHierarchyNode;
    update: (source: CustomHierarchyNode) => void;
    zoom: d3.ZoomBehavior<SVGSVGElement, unknown>;
  }>();

  // Add collapse all functionality
  const collapseAll = (): void => {
    if (d3Ref.current) {
      const { root, update, zoom } = d3Ref.current;

      const collapse = (node: CustomHierarchyNode): void => {
        if (node.children) {
          node._children = node.children;
          node.children = undefined;
        }
        if (node._children) {
          node._children.forEach(collapse);
        }
      };

      collapse(root);
      update(root);

      // Reset view to root
      const initialTransform = d3.zoomIdentity.translate(containerWidth / 2, 80).scale(1);
      d3.select<SVGSVGElement, unknown>('#treeSvg')
        .transition()
        .duration(750)
        .call(zoom.transform, initialTransform);
    }
  };

  const expandAll = (): void => {
    if (d3Ref.current) {
      const { root, update, zoom } = d3Ref.current;

      const expand = (node: CustomHierarchyNode): void => {
        if (node._children) {
          node.children = node._children;
          node._children = undefined;
        }
        if (node.children) {
          node.children.forEach(expand);
        }
      };

      expand(root);
      update(root);

      // Reset view to root
      const initialTransform = d3.zoomIdentity.translate(containerWidth / 2, 80).scale(1);
      d3.select<SVGSVGElement, unknown>('#treeSvg')
        .transition()
        .duration(750)
        .call(zoom.transform, initialTransform);
    }
  };

  // In the useEffect where d3Ref is set
  useEffect(() => {
    if (treeData.length > 0) {
      const svgWidth = containerWidth;
      const nodeRadius = 12;
      const width = containerWidth;
      const height = 600;

      const svg = d3.select<SVGSVGElement, unknown>('#treeSvg')
        .attr('width', width)
        .attr('height', height);

      const tooltip = d3.select<HTMLDivElement, unknown>('#tooltip');

      const root = d3.hierarchy(treeData[0]) as CustomHierarchyNode;

      // Modified initial node setup
      root.descendants().forEach((d) => {
        const node = d as CustomHierarchyNode;
        if (node !== root && node.children) {
          node._children = node.children;
          node.children = undefined;
        }
      });

      // Move g definition above zoom to fix use-before-define
      const g = svg.append('g');

      const zoom = d3.zoom<SVGSVGElement, unknown>()
        .scaleExtent([0.5, 3])
        .on('zoom', (event) => {
          g.attr('transform', event.transform);
        });

      svg.call(zoom);

      const treeLayout = d3.tree<D3Node>()
        .nodeSize([100, 200])
        .separation((a, b) => (a.depth === 0 || b.depth === 0 ? 1 : 2));

      let tree = treeLayout(root);

      const update = (source: CustomHierarchyNode): void => {
        tree = treeLayout(root);

        const nodes = tree.descendants();
        const links = tree.links();

        const link = g.selectAll<SVGPathElement, unknown>('.link')
          .data(links, (d: { source: CustomHierarchyNode; target: CustomHierarchyNode }) => `${d.source.data.title}-${d.target.data.title}`);

        link.enter()
          .append('path')
          .attr('class', 'link')
          .merge(link)
          .transition()
          .duration(750)
          .attr('d', (d: { source: CustomHierarchyNode; target: CustomHierarchyNode }) => {
            const diagonal = d3.linkVertical<{ source: CustomHierarchyNode; target: CustomHierarchyNode }, unknown>()
              .x((d) => (d as CustomHierarchyNode).x ?? 0)
              .y((d) => (d as CustomHierarchyNode).y ?? 0);
            return diagonal(d) || '';
          })
          .attr('fill', 'none')
          .attr('stroke', '#ccc')
          .attr('stroke-width', 2);

        link.exit().remove();

        const node = g.selectAll<SVGGElement, unknown>('.node')
          .data(nodes, (d: CustomHierarchyNode) => d.data.title);

        const nodeEnter = node.enter()
          .append('g')
          .attr('class', 'node')
          .attr('transform', (d: CustomHierarchyNode) => `translate(${d.x},${d.y})`)
          .on('click', (event, d: CustomHierarchyNode) => {
            if (d.children) {
              d._children = d.children;
              d.children = undefined;
            } else {
              d.children = d._children;
              d._children = undefined;
            }
            update(d);
          })
          .on('mouseover', (event, d: CustomHierarchyNode) => {
            let tooltipHtml = `<strong>${d.data.title}</strong>`;
            
            const levelTooltips = props.tooltipFields[d.data.level] || [];
            levelTooltips.forEach(field => {
              if (d.data.tooltipData?.[field]) {
                tooltipHtml += `<br><strong>${props.columnDisplayNames[field]}:</strong> ${d.data.tooltipData[field]}`;
              }
            });

            tooltip.style('visibility', 'visible')
              .style('top', `${event.pageY + 10}px`)
              .style('left', `${event.pageX + 10}px`)
              .html(tooltipHtml);
          })
          .on('mousemove', (event) => {
            tooltip.style('top', `${event.pageY + 10}px`)
              .style('left', `${event.pageX + 10}px`);
          })
          .on('mouseout', () => {
            tooltip.style('visibility', 'hidden');
          });

        nodeEnter.append('circle')
          .attr('r', nodeRadius)
          .attr('fill', '#0078d4')
          .on('mouseover', function () {
            d3.select(this).attr('fill', '#68c1e8');
          })
          .on('mouseout', function () {
            d3.select(this).attr('fill', '#0078d4');
          });

        nodeEnter.append('foreignObject')
          .attr('x', 20)
          .attr('y', -10)
          .attr('width', 200)
          .attr('height', 40)
          .append('xhtml:div')
          .style('font-size', '12px')
          .style('word-wrap', 'break-word')
          .style('white-space', 'normal')
          .style('max-width', '180px')
          .text((d: CustomHierarchyNode) => d.data.title);

        const mergedNodes = nodeEnter.merge(node);

        mergedNodes.transition()
          .duration(750)
          .attr('transform', (d: CustomHierarchyNode) => `translate(${d.x},${d.y})`);

        mergedNodes.select('circle')
          .transition()
          .duration(750)
          .attr('fill', (d: CustomHierarchyNode) => {
            const level = d.data.level;
            return props.nodeColors[level] || '#0078d4'; // Fallback to default color
          });

        node.exit().remove();
      };

      update(root);

      const initialTransform = d3.zoomIdentity.translate(svgWidth / 2, 80).scale(1);
      svg.call(zoom.transform, initialTransform);
      d3Ref.current = { root, update, zoom }; // Store zoom in ref

    }
  }, [treeData, props.tooltipFields, containerWidth]);

  return (
    <div ref={containerRef} className={styles.treeContainer} style={{ position: 'relative' }}>
      {/* Expand All Button */}
      <button
        className={styles.controlButton}
        onClick={expandAll}
        title="Expand All"
        style={{
          position: 'absolute',
          top: '10px',
          left: '10px',
          padding: '5px',
          backgroundColor: 'white',
          border: '1px solid #ccc',
          fontSize: '18px',
          borderRadius: '4px',
          color: 'black',
          cursor: 'pointer',
          outline: 'none',
        }}
      >
        <FontAwesomeIcon icon={faExpand} size="sm" />
      </button>
  
      {/* Collapse All Button */}
      <button
        className={styles.controlButton}
        onClick={collapseAll}
        title="Collapse All"
        style={{
          position: 'absolute',
          top: '10px',
          left: '40px',
          padding: '5px',
          fontSize: '18px',
          border: '1px solid #ccc',
          borderRadius: '4px',
          backgroundColor: 'white',
          color: 'black',
          cursor: 'pointer',
          outline: 'none',
        }}
      >
        <FontAwesomeIcon icon={faCompress} size="sm" />
      </button>
      
      <svg id="treeSvg" />
      <div id="tooltip" style={{ position: 'fixed', top: '10px', left: '10px', backgroundColor: 'white', border: '1px solid #ccc', padding: '10px', borderRadius: '5px', boxShadow: '0 0 10px rgba(0,0,0,0.1)', visibility: 'hidden' }} />
    </div>
  );
};
export default TreeWebPart;