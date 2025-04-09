import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import { convert } from "html-to-text";
import * as fs from "fs";
import { config } from "dotenv";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListPromptsRequestSchema,
  GetPromptRequestSchema,
  Prompt,
  Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import { PublicClientApplication, DeviceCodeRequest } from "@azure/msal-node";

// Load environment variables from .env.local
const envPath = join(dirname(fileURLToPath(import.meta.url)), "../.env.local");
if (fs.existsSync(envPath)) {
  config({ path: envPath });
} else {
  console.warn("Warning: .env.local file not found. Using environment variables from system.");
}

// Validate required environment variables
if (!process.env.CLIENT_ID) {
  throw new Error(
    "CLIENT_ID environment variable is required. Please create a .env.local file with your Azure client ID or set it in your environment."
  );
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Define the prompts
const getNotebooksPrompt: Prompt = {
  id: "get-notebooks",
  name: "Get all notebooks",
  description: "List all available OneNote notebooks",
  text: "Show me a list of all my OneNote notebooks.",
};

const getSectionsPrompt: Prompt = {
  id: "get-sections",
  name: "Get sections in a notebook",
  description: "List all sections in a specific notebook",
  text: 'Show me all sections in my notebook called "[NOTEBOOK_NAME]". To do this, you\'ll need to:\n1. First get a list of all my notebooks\n2. Find "[NOTEBOOK_NAME]" in the list and note its ID\n3. Use that ID to get all sections in the notebook',
};

const getPagesPrompt: Prompt = {
  id: "get-pages",
  name: "Get pages in a section",
  description: "List all pages in a specific section",
  text: 'Show me all pages in the section "[SECTION_NAME]" of my notebook "[NOTEBOOK_NAME]". To do this, you\'ll need to:\n1. First get a list of all my notebooks\n2. Find "[NOTEBOOK_NAME]" in the list and note its ID\n3. Use that ID to get all sections in the notebook\n4. Find "[SECTION_NAME]" in the list and note its ID\n5. Use that section ID to get all pages in the section',
};

const getPagePrompt: Prompt = {
  id: "get-page",
  name: "Get a specific page",
  description: "Get content from a specific page",
  text: 'Show me the content of the page "[PAGE_NAME]" in the section "[SECTION_NAME]" of my notebook "[NOTEBOOK_NAME]". To do this, you\'ll need to:\n1. First get a list of all my notebooks\n2. Find "[NOTEBOOK_NAME]" in the list and note its ID\n3. Use that ID to get all sections in the notebook\n4. Find "[SECTION_NAME]" in the list and note its ID\n5. Use that section ID to get all pages in the section\n6. Find "[PAGE_NAME]" in the list and note its ID\n7. Use that page ID to get the content of the page',
};

const createNotebookPrompt: Prompt = {
  id: "create-notebook",
  name: "Create a new notebook",
  description: "Create a brand new notebook",
  text: 'Create a new OneNote notebook called "[NOTEBOOK_NAME]".',
};

const createSectionPrompt: Prompt = {
  id: "create-section",
  name: "Create a section in a notebook",
  description: "Create a new section in a specific notebook",
  text: 'I want you to create a section called "[SECTION_NAME]" in my notebook called "[NOTEBOOK_NAME]". To do this, you\'ll need to:\n1. First get a list of all my notebooks\n2. Find "[NOTEBOOK_NAME]" in the list and note its ID\n3. Use that ID to create the new section',
};

const createPagePrompt: Prompt = {
  id: "create-page",
  name: "Create a page in a section",
  description: "Create a new page in a specific section",
  text: 'I want you to create a page called "[PAGE_NAME]" in the section "[SECTION_NAME]" of my notebook "[NOTEBOOK_NAME]". To do this, you\'ll need to:\n1. First get a list of all my notebooks\n2. Find "[NOTEBOOK_NAME]" in the list and note its ID\n3. Use that ID to get all sections in the notebook\n4. Find "[SECTION_NAME]" in the list and note its ID\n5. Use that section ID to create the new page',
};

// Store all prompts in an array
const prompts = [
  getNotebooksPrompt,
  getSectionsPrompt,
  getPagesPrompt,
  getPagePrompt,
  createNotebookPrompt,
  createSectionPrompt,
  createPagePrompt,
];

// OneNote tool definitions
const oneNoteReadTool: Tool = {
  name: "onenote-read",
  description:
    "Read content from Microsoft OneNote notebooks, sections, or pages",
  inputSchema: {
    type: "object",
    properties: {
      type: {
        type: "string",
        enum: ["read_content"],
      },
      pageId: {
        type: "string",
        description: "ID of the specific page to read",
      },
      sectionId: {
        type: "string",
        description: "ID of the section to list pages from",
      },
      notebookId: {
        type: "string",
        description: "ID of the notebook to list sections from",
      },
      sectionGroupId: {
        type: "string",
        description: "ID of the section group to list sections from",
      },
      includeContent: {
        type: "boolean",
        default: true,
        description: "Whether to include the content of the page",
      },
      includeMetadata: {
        type: "boolean",
        default: false,
        description: "Whether to include metadata about the page",
      },
    },
    required: ["type"],
  },
};

const oneNoteCreateTool: Tool = {
  name: "onenote-create",
  description: "Create new content in Microsoft OneNote",
  inputSchema: {
    type: "object",
    properties: {
      type: {
        type: "string",
        enum: ["create_page", "create_section", "create_notebook"],
      },
      title: {
        type: "string",
        description: "Title of the content to create",
      },
      content: {
        type: "string",
        description: "Content in Markdown format",
      },
      parentId: {
        type: "string",
        description: "ID of the parent section or notebook",
      },
    },
    required: ["type", "content"],
  },
};

// OneNote service class
// This class handles the interaction with Microsoft Graph API for OneNote
// and manages authentication using MSAL
class OneNoteService {
  private client: Client | null = null;
  private tokenCache: any = null;
  private pca: PublicClientApplication;
  private readonly cacheFile = join(__dirname, "token-cache.json");

  constructor(
    private config: {
      clientId: string;
    }
  ) {
    // Setup token cache persistence
    const beforeCacheAccess = async (cacheContext: any) => {
      if (fs.existsSync(this.cacheFile)) {
        cacheContext.tokenCache.deserialize(
          fs.readFileSync(this.cacheFile, "utf-8")
        );
      }
    };

    const afterCacheAccess = async (cacheContext: any) => {
      if (cacheContext.cacheHasChanged) {
        fs.writeFileSync(
          this.cacheFile,
          cacheContext.tokenCache.serialize(),
          "utf-8"
        );
      }
    };

    const cachePlugin = {
      beforeCacheAccess,
      afterCacheAccess,
    };

    const msalConfig = {
      auth: {
        clientId: this.config.clientId,
        authority: "https://login.microsoftonline.com/common",
      },
      cache: {
        cachePlugin,
      },
    };

    this.pca = new PublicClientApplication(msalConfig);
  }

  async initialize() {
    this.tokenCache = this.pca.getTokenCache();

    // Request permissions for OneNote
    const scopes = ["Notes.Read", "Notes.ReadWrite"];
    try {
      // Try to get token silently first
      let authResult;
      const accounts = await this.tokenCache.getAllAccounts();

      if (accounts.length > 0) {
        try {
          authResult = await this.pca.acquireTokenSilent({
            scopes,
            account: accounts[0],
          });
        } catch (silentError) {
          // Fall back to device code flow
          const deviceCodeRequest: DeviceCodeRequest = {
            deviceCodeCallback: async (response) => {
              await this.addInstructionsToFile(response);
            },
            scopes,
          };

          authResult = await this.pca.acquireTokenByDeviceCode(
            deviceCodeRequest
          );
        }
      } else {
        // No accounts in cache, use device code flow
        const deviceCodeRequest: DeviceCodeRequest = {
          deviceCodeCallback: async (response) => {
            // This will display the device code and instructions
            await this.addInstructionsToFile(response);
          },
          scopes,
        };

        authResult = await this.pca.acquireTokenByDeviceCode(deviceCodeRequest);
      }

      // Initialize Graph client
      this.client = Client.init({
        authProvider: (done) => {
          if (authResult) {
            done(null, authResult.accessToken);
          } else {
            done(new Error("Failed to acquire token"), null);
          }
        },
      });

      return this;
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(
          `Failed to initialize OneNote service: ${error.message}`
        );
      } else {
        throw new Error("Failed to initialize OneNote service: Unknown error");
      }
    }
  }
  async addInstructionsToFile(response: any) {
    // This will display the device code and instructions
    const deviceCodePath = join(__dirname, "device-code.txt");
    const content = `
  -----------------------------------------------------
  MICROSOFT AUTHENTICATION REQUIRED
  -----------------------------------------------------
  ${response.message}
  
  This file was generated at ${new Date().toISOString()}
  -----------------------------------------------------
  
  `;
    // Write the instructions to a file
    fs.writeFileSync(deviceCodePath, content);
  }

  async getNotebooks() {
    const response = await this.client!.api(`/users/me/onenote/notebooks`)
      .select("id,displayName,createdDateTime,lastModifiedDateTime")
      .get();
    return response.value;
  }

  async getSections(notebookId: string) {
    const response = await this.client!.api(
      `/me/onenote/notebooks/${notebookId}/sections`
    )
      .select("id,displayName,pagesUrl")
      .get();
    return response.value;
  }

  async getSectionGroups(notebookId: string) {
    const response = await this.client!.api(
      `/me/onenote/notebooks/${notebookId}/sectionGroups`
    )
      .select("id,displayName,sectionsUrl,sectionGroupsUrl")
      .get();
    return response.value;
  }

  async getSectionsFromGroup(sectionGroupId: string) {
    const response = await this.client!.api(
      `/me/onenote/sectionGroups/${sectionGroupId}/sections`
    )
      .select("id,displayName,pagesUrl")
      .get();
    return response.value;
  }

  async getPages(sectionId: string) {
    const response = await this.client!.api(
      `/me/onenote/sections/${sectionId}/pages`
    )
      .select("id,title,createdDateTime,lastModifiedDateTime")
      .get();
    return response.value;
  }

  async getPage(pageId: string, includeContent: boolean = true) {
    let page = await this.client!.api(`/me/onenote/pages/${pageId}`)
      .select("id,title,createdDateTime,lastModifiedDateTime,level")
      .get();

    if (includeContent) {
      const contentResponse = await this.client!.api(
        `/me/onenote/pages/${pageId}/content`
      )
        .header("Accept", "text/html")
        .responseType(ResponseType.TEXT)
        .get();

      page.content = contentResponse;

      page.textContent = convert(contentResponse, {
        wordwrap: null,
        selectors: [
          { selector: "a", options: { ignoreHref: false } },
          { selector: "img", format: "skip" },
          { selector: "table", format: "dataTable" },
        ],
      });
    }

    return page;
  }

  async createPage(sectionId: string, title: string, htmlContent: string) {
    // Directly use the HTML content provided by the agent
    const response = await this.client!.api(
      `/me/onenote/sections/${sectionId}/pages`
    )
      .header("Content-Type", "application/xhtml+xml")
      .post(htmlContent);

    return response;
  }

  async createSection(notebookId: string, name: string) {
    const response = await this.client!.api(
      `/me/onenote/notebooks/${notebookId}/sections`
    ).post({
      displayName: name,
    });

    return response;
  }

  async createNotebook(name: string) {
    const response = await this.client!.api("/me/onenote/notebooks").post({
      displayName: name,
    });

    return response;
  }
}

// Create MCP Server
async function main() {
  // Initialize OneNote service
  const oneNoteService = await new OneNoteService({
    clientId: process.env.CLIENT_ID || "",
  }).initialize();

  const server = new Server({
    name: "onenote-mcp-server",
    version: "1.0.0",
    capabilities: {
      tools: {
        oneNoteReadTool,
        oneNoteCreateTool,
      },
    },
  });

  // Handle list prompts request
  server.setRequestHandler(ListPromptsRequestSchema, async (request) => {
    return {
      prompts: prompts.map((prompt) => ({
        id: prompt.id,
        name: prompt.name,
        description: prompt.description,
      })),
    };
  });

  // Handle get prompt request
  server.setRequestHandler(GetPromptRequestSchema, async (request) => {
    const { params } = request;
    const { id } = params;

    const prompt = prompts.find((p) => p.id === id);
    if (!prompt) {
      throw new Error(`Prompt not found: ${id}`);
    }

    return {
      prompt,
    };
  });
  
  // Handle list tools request
  server.setRequestHandler(ListToolsRequestSchema, async (request) => {
    return {
      tools: [oneNoteReadTool, oneNoteCreateTool],
    };
  });

  // Handle tool calls
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { params } = request;
    const { name, arguments: parameters = {} } = params;

    try {
      // Handle OneNote read operations
      if (name === "onenote-read") {
        const type = parameters.type as string;
        const pageId = parameters.pageId as string;
        const sectionId = parameters.sectionId as string;
        const notebookId = parameters.notebookId as string;
        const sectionGroupId = parameters.sectionGroupId as string;
        const includeContent = parameters.includeContent as boolean;
        const includeMetadata = parameters.includeMetadata as boolean;

        if (pageId) {
          const page = await oneNoteService.getPage(pageId, includeContent);
          const result: any = {};

          if (includeContent) {
            // Pass the content directly as retrieved from OneNote
            result.content = page.content;
            result.textContent = page.textContent;
          }

          if (includeMetadata) {
            result.metadata = {
              id: page.id,
              title: page.title,
              createdTime: page.createdDateTime,
              lastModifiedTime: page.lastModifiedDateTime,
            };
          }

          return {
            content: [
              {
                type: "text",
                text: JSON.stringify(result),
              },
            ],
            isError: false,
          };
        } else if (sectionId) {
          const pages = await oneNoteService.getPages(sectionId);
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  sectionId,
                  pages: pages.map((page: any) => ({
                    id: page.id,
                    title: page.title,
                    lastModifiedTime: page.lastModifiedDateTime,
                  })),
                }),
              },
            ],
            isError: false,
          };
        } else if (sectionGroupId) {
          const sections = await oneNoteService.getSectionsFromGroup(sectionGroupId);
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  sectionGroupId,
                  sections: sections.map((section: any) => ({
                    id: section.id,
                    name: section.displayName,
                  })),
                }),
              },
            ],
            isError: false,
          };
        } else if (notebookId) {
          const [sections, sectionGroups] = await Promise.all([
            oneNoteService.getSections(notebookId),
            oneNoteService.getSectionGroups(notebookId)
          ]);
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  notebookId,
                  sections: sections.map((section: any) => ({
                    id: section.id,
                    name: section.displayName,
                  })),
                  sectionGroups: sectionGroups.map((group: any) => ({
                    id: group.id,
                    name: group.displayName,
                  })),
                }),
              },
            ],
            isError: false,
          };
        } else {
          const notebooks = await oneNoteService.getNotebooks();
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  notebooks: notebooks.map((nb: any) => ({
                    id: nb.id,
                    name: nb.displayName,
                  })),
                }),
              },
            ],
            isError: false,
          };
        }
      }

      // Handle OneNote create operations
      if (name === "onenote-create") {
        const type = parameters.type as string;
        const title = parameters.title as string;
        const content = parameters.content as string;
        const parentId = parameters.parentId as string;

        if (type === "create_page" && parentId) {
          const newPage = await oneNoteService.createPage(
            parentId,
            title as string,
            content as string
          );
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  id: newPage.id,
                  title: newPage.title,
                  createdTime: newPage.createdDateTime,
                }),
              },
            ],
            isError: false,
          };
        } else if (type === "create_section" && parentId) {
          const newSection = await oneNoteService.createSection(
            parentId as string,
            title as string
          );
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  id: newSection.id,
                  name: newSection.displayName,
                }),
              },
            ],
            isError: false,
          };
        } else if (type === "create_notebook") {
          const newNotebook = await oneNoteService.createNotebook(
            title as string
          );
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  id: newNotebook.id,
                  name: newNotebook.displayName,
                }),
              },
            ],
            isError: false,
          };
        }
      }

      throw new Error(`Unsupported tool or operation: ${name}`);
    } catch (error: unknown) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      throw new Error(`Failed to execute tool: ${errorMessage}`);
    }
  });

  // Start the server with stdio transport
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

// Run the server
main().catch((error) => {
  process.exit(1);
});
