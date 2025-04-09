# OneNote MCP Server

A Model Context Protocol (MCP) server that provides AI assistants with access to Microsoft OneNote. This server enables AI models to read from and write to OneNote notebooks, sections, and pages.

## Project Overview

This project implements an MCP server that connects to Microsoft OneNote using the Microsoft Graph API. It provides tools for:

- Reading notebooks, sections, and pages from OneNote
- Creating new notebooks, sections, and pages in OneNote
- Converting HTML content to text for better RAG processing

![Demo](Demo.gif)

## Project Structure

```
onenote/
├── dist/                # Compiled JavaScript files (generated)
├── src/                 # TypeScript source files
│   ├── index.ts         # Main entry point and server implementation
│   └── types/           # Custom TypeScript type definitions
├── .vscode/             # VS Code configuration
│   ├── launch.json      # Debug configurations
│   └── tasks.json       # Build tasks
├── package.json         # Project dependencies and scripts
├── tsconfig.json        # TypeScript configuration
├── Dockerfile           # Docker configuration
├── .env.local.example   # Example environment variables
└── README.md            # This file
```

## Authentication

The server uses Microsoft Authentication Library (MSAL) with device code flow for authentication:

1. When first run, the server generates a device code and URL
2. The code is saved to `device-code.txt` in the project directory
3. You must visit the URL and enter the code to authenticate
4. After authentication, tokens are cached in `token-cache.json` for future use

## MCP Tools

The server provides the following MCP tools:

### onenote-read

Read content from Microsoft OneNote notebooks, sections, or pages.

Parameters:

- `type`: "read_content"
- `pageId`: (optional) ID of the specific page to read
- `sectionId`: (optional) ID of the section to list pages from
- `notebookId`: (optional) ID of the notebook to list sections from
- `includeContent`: (optional) Whether to include the content of the page (default: true)
- `includeMetadata`: (optional) Whether to include metadata about the page (default: false)

### onenote-create

Create new content in Microsoft OneNote.

Parameters:

- `type`: "create_page", "create_section", or "create_notebook"
- `title`: Title of the content to create
- `content`: Content in HTML format (for pages)
- `parentId`: (optional) ID of the parent section or notebook

## Getting Started

### Prerequisites

- Node.js (v14 or higher)
- npm (v6 or higher)
- Microsoft Azure account with a registered application
- OneNote account (Microsoft 365 subscription)

### Azure Setup

1. Register a new application in the [Azure Portal](https://portal.azure.com)
2. Add the following API permissions:
   - Microsoft Graph > Notes.Read
   - Microsoft Graph > Notes.ReadWrite
3. Configure authentication:
   - Under "Authentication" settings, set "Supported account types" to:
     - "Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
   - Enable "Allow public client flows" for the app
4. Note your Application (client) ID for configuration

### Installation

1. Clone the repository
2. Install dependencies:

```bash
npm install
```

3. Create a `.env.local` file with your Azure client ID:

```
CLIENT_ID=your-client-id-from-azure
```

### Development

To run the application in development mode:

```bash
npm run dev
```

### Building

To compile TypeScript to JavaScript:

```bash
npm run build
```

### Running

To run the compiled application:

```bash
npm start
```

## Docker Support

You can build and run the application using Docker:

```bash
# Create a data directory for persistence
mkdir -p data

# Build the Docker image
docker build -t onenote-mcp-server .

# Run the container
docker run -d \
  --name onenote-mcp-server \
  -e CLIENT_ID=your-client-id \
  -v $(pwd)/data:/app/dist \
  onenote-mcp-server
```

### Authentication with Docker

When running in Docker, the authentication flow works as follows:

1. Start the container as shown above
2. Check the device code file:
   ```bash
   cat data/device-code.txt
   ```
3. Follow the instructions to authenticate with Microsoft
4. The token will be cached in `data/token-cache.json` for future use

## Setup with Claude Desktop

1. Clone this repository
2. Run `npm install` to install dependencies
3. In Claude Desktop, add a new MCP server:
   - Set the server directory to your cloned repository
   - Set the command to: `npm run build && npm start`
   - Add the following environment variable:
     - Name: CLIENT_ID
     - Value: [Your Microsoft Azure Application Client ID]
4. Save the configuration and connect to the server

## Authentication Flow

1. On first run, the server will generate a device code and URL
2. The code is saved to `device-code.txt` in the project directory
3. Visit the URL and enter the code to authenticate
4. After authentication, tokens are cached in `token-cache.json`
5. Subsequent runs will use the cached token if valid

## Troubleshooting

- **Authentication Issues**: Delete `token-cache.json` to force re-authentication
- **Module Errors**: Ensure you're using Node.js 14+ with ES modules support
- **TypeScript Errors**: Run `npm run build` to check for compilation errors
