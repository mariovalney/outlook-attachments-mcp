/**
 * Folder management module for Outlook Assistant server
 */
const handleListFolders = require('./list');
const handleCreateFolder = require('./create');
const handleMoveEmails = require('./move');
const handleGetFolderStats = require('./stats');
const handleDeleteFolder = require('./delete');

// Consolidated folder tool definition
const folderTools = [
  {
    name: 'folders',
    description:
      'Manage mail folders (tool-level destructiveHint=true because `delete` permanently removes a folder; `list` and `stats` are read-only sub-actions despite the annotation). action=`list` (default) returns the folder tree with id/displayName/parentFolderId (toggle `includeItemCounts` for unread/total, `includeChildren` for hierarchy). action=`create` makes a new folder under the inbox (or under `folder`/`folderId`/`folderName`) and returns its id. action=`move` relocates emails (`emailIds` array) into `targetFolder`. action=`stats` returns counts (totalItemCount/unreadItemCount) suitable for pagination planning — pair with `outputVerbosity` to limit noise. action=`delete` permanently removes a folder and its contents — there is no recycle-bin recovery.',
    annotations: {
      title: 'Mail Folders',
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: false,
    },
    inputSchema: {
      type: 'object',
      properties: {
        action: {
          type: 'string',
          enum: ['list', 'create', 'move', 'stats', 'delete'],
          description: 'Action to perform (default: list)',
        },
        // list params
        includeItemCounts: {
          type: 'boolean',
          description: 'Include counts of total and unread items (action=list)',
        },
        includeChildren: {
          type: 'boolean',
          description: 'Include child folders in hierarchy (action=list)',
        },
        // create params
        name: {
          type: 'string',
          description: 'Name of the folder to create (action=create, required)',
        },
        parentFolder: {
          type: 'string',
          description: 'Parent folder name, default is root (action=create)',
        },
        // move params
        emailIds: {
          type: 'string',
          description:
            'Comma-separated list of email IDs to move (action=move, required)',
        },
        targetFolder: {
          type: 'string',
          description: 'Folder name to move emails to (action=move, required)',
        },
        sourceFolder: {
          type: 'string',
          description: 'Source folder name, default is inbox (action=move)',
        },
        // stats params
        folder: {
          type: 'string',
          description:
            'Folder name (inbox, sent, drafts, etc.). Default: inbox (action=stats)',
        },
        outputVerbosity: {
          type: 'string',
          enum: ['minimal', 'standard', 'full'],
          description: 'Output detail level (action=stats, default: standard)',
        },
        // delete params
        folderId: {
          type: 'string',
          description: 'Folder ID to delete (action=delete)',
        },
        folderName: {
          type: 'string',
          description:
            'Folder name to delete — resolved to ID (action=delete). Cannot delete protected folders (Inbox, Drafts, Sent, etc.)',
        },
      },
      additionalProperties: false,
      required: [],
    },
    handler: async (args) => {
      const action = args.action || 'list';
      switch (action) {
        case 'create':
          return handleCreateFolder(args);
        case 'move':
          return handleMoveEmails(args);
        case 'stats':
          return handleGetFolderStats(args);
        case 'delete':
          return handleDeleteFolder(args);
        case 'list':
          return handleListFolders(args);
        default:
          return {
            content: [
              {
                type: 'text',
                text: `Unknown action '${action}'. Valid actions: list, create, move, stats, delete.`,
              },
            ],
          };
      }
    },
  },
];

module.exports = {
  folderTools,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails,
  handleGetFolderStats,
  handleDeleteFolder,
};
