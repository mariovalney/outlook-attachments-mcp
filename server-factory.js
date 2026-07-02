/**
 * MCP Server factory — shared between the stdio entry point (index.js)
 * and the HTTP entry point (http-server.js).
 *
 * Builds the tool list and a configured Server instance. The HTTP mode
 * omits the device-code `auth` tool: authentication there happens at the
 * connector level (OAuth with Microsoft), so every request already carries
 * the user's Graph token.
 */
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const config = require('./config');
const { coerceArgsAgainstSchema } = require('./utils/schema-coerce');

// Import module tools
const { authTools, setToolCount } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { contactsTools } = require('./contacts');
const { categoriesTools } = require('./categories');
const { settingsTools } = require('./settings');
const { advancedTools } = require('./advanced');

/**
 * Builds the combined tool list.
 * @param {object} options
 * @param {boolean} options.includeAuthTools - Include the device-code `auth`
 *   tool (stdio/single-user mode only).
 * @returns {Array} tools
 */
function buildTools({ includeAuthTools = true } = {}) {
  const tools = [
    ...(includeAuthTools ? authTools : []),
    ...calendarTools,
    ...emailTools,
    ...folderTools,
    ...rulesTools,
    ...contactsTools,
    ...categoriesTools,
    ...settingsTools,
    ...advancedTools,
  ];

  // Set dynamic tool count for auth about handler
  setToolCount(tools.length);

  return tools;
}

/**
 * Creates a configured MCP Server wired to the given tool list.
 * @param {Array} TOOLS - Tool list from buildTools()
 * @returns {Server}
 */
function createServer(TOOLS) {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    {
      capabilities: {
        tools: TOOLS.reduce((acc, tool) => {
          acc[tool.name] = {};
          return acc;
        }, {}),
      },
    }
  );

  // Handle all requests
  server.fallbackRequestHandler = async (request) => {
    try {
      const { method, params, id } = request;
      console.error(`REQUEST: ${method} [${id}]`);

      // Initialize handler
      if (method === 'initialize') {
        console.error(`INITIALIZE REQUEST: ID [${id}]`);
        return {
          protocolVersion: '2024-11-05',
          capabilities: {
            tools: TOOLS.reduce((acc, tool) => {
              acc[tool.name] = {};
              return acc;
            }, {}),
          },
          serverInfo: {
            name: config.SERVER_NAME,
            version: config.SERVER_VERSION,
          },
        };
      }

      // Tools list handler
      if (method === 'tools/list') {
        console.error(`TOOLS LIST REQUEST: ID [${id}]`);
        console.error(`TOOLS COUNT: ${TOOLS.length}`);

        return {
          tools: TOOLS.map((tool) => ({
            name: tool.name,
            description: tool.description,
            inputSchema: tool.inputSchema,
            ...(tool.annotations && { annotations: tool.annotations }),
          })),
        };
      }

      // Required empty responses for other capabilities
      if (method === 'resources/list') return { resources: [] };
      if (method === 'prompts/list') return { prompts: [] };

      // Tool call handler
      if (method === 'tools/call') {
        try {
          const { name, arguments: args = {} } = params || {};

          console.error(`TOOL CALL: ${name}`);

          // Find the tool handler
          const tool = TOOLS.find((t) => t.name === name);

          if (tool && tool.handler) {
            // Coerce + validate args against the tool's inputSchema before
            // dispatching. Catches array-as-string, boolean-as-string, unknown
            // params, and out-of-enum action values at the MCP boundary so
            // handlers receive properly-typed JS values. (#160, #162)
            if (tool.inputSchema) {
              const coerced = coerceArgsAgainstSchema(args, tool.inputSchema);
              if (coerced.error) {
                return {
                  content: [
                    {
                      type: 'text',
                      text: `Invalid arguments for tool '${name}':\n${coerced.error}`,
                    },
                  ],
                  isError: true,
                };
              }
              return await tool.handler(coerced.args);
            }
            return await tool.handler(args);
          }

          // Tool not found
          return {
            error: {
              code: -32601,
              message: `Tool not found: ${name}`,
            },
          };
        } catch (error) {
          console.error(`Error in tools/call:`, error);
          return {
            error: {
              code: -32603,
              message: `Error processing tool call: ${error.message}`,
            },
          };
        }
      }

      // For any other method, return method not found
      return {
        error: {
          code: -32601,
          message: `Method not found: ${method}`,
        },
      };
    } catch (error) {
      console.error(`Error in fallbackRequestHandler:`, error);
      return {
        error: {
          code: -32603,
          message: `Error processing request: ${error.message}`,
        },
      };
    }
  };

  return server;
}

module.exports = { buildTools, createServer };
