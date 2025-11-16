# SharePoint D3 Tree WebPart

## Summary

A powerful SharePoint Framework (SPFx) WebPart that creates interactive hierarchical tree visualizations using D3.js. This solution allows users to build dynamic tree structures from any SharePoint list data with customizable styling, filtering, and multi-level tooltips.

### Key Features

- **Dynamic Tree Visualization**: Creates interactive D3.js tree diagrams from SharePoint list data
- **Indexed Lookup Support**: Handles SharePoint indexed lookup columns and related fields
- **Advanced Filtering**: Supports OR/AND filtering with multiple values for precise data display
- **Multi-Level Tooltips**: Configurable tooltips for each tree level showing relevant field data
- **Customizable Styling**: Level-based color customization for tree nodes
- **Mobile Responsive**: Touch-friendly controls with expand/collapse functionality
- **Real-time Data**: Fetches live data from SharePoint lists using PnP/SP

![SharePoint D3 Tree WebPart](https://img.shields.io/badge/SharePoint-Framework-blue) ![D3.js](https://img.shields.io/badge/D3.js-Tree-orange) ![TypeScript](https://img.shields.io/badge/TypeScript-Ready-green)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| SharePoint D3 Tree WebPart | Karim Hassan (kareemelthird) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | November 16, 2025 | Initial release with D3 tree visualization, indexed lookup support, and advanced filtering |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

This SharePoint Framework WebPart provides comprehensive tree visualization capabilities with the following advanced features:

### Tree Building and Visualization
- **Hierarchical Data Structure**: Automatically builds tree hierarchies from SharePoint list data
- **Interactive D3.js Visualization**: Smooth animations, zoom/pan functionality, and responsive design
- **Multi-Level Support**: Supports up to 7 levels of tree hierarchy
- **Expand/Collapse Controls**: User-friendly buttons to expand or collapse all nodes

### SharePoint Integration
- **PnP/SP Integration**: Uses @pnp/sp for efficient SharePoint REST API calls
- **Indexed Lookup Fields**: Full support for SharePoint indexed lookup columns and projected fields
- **Related Field Handling**: Properly processes lookup relationships and related data
- **Real-time Data Fetching**: Retrieves live data with pagination support for large datasets

### Advanced Filtering System
- **Multi-Value Filtering**: Filter by multiple values using OR/AND logic operators
- **Flexible Filter Configuration**: Choose any column for filtering with custom text values
- **Dynamic Query Building**: Constructs efficient OData filters for optimal performance

### Customization Options
- **Level-Based Colors**: Customize node colors for each hierarchy level
- **Multi-Field Tooltips**: Configure up to 3 tooltip fields per tree level
- **Responsive Design**: Automatically adapts to container width and mobile devices
- **Property Pane Controls**: Rich configuration interface with dropdowns and color pickers

This solution demonstrates advanced concepts including:
- Complex SharePoint field metadata processing
- D3.js tree layout and interaction handling
- TypeScript interfaces and React functional components
- SPFx property pane customization and validation
- Efficient data fetching and caching strategies

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
