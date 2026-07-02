/**
 * Attachment handling for Outlook Assistant server
 * Provides tools to list and download email attachments via Microsoft Graph API
 */
const _https = require('https'); // Reserved for future use
const _config = require('../config'); // Reserved for future use
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

// This is a remote, multi-user MCP server: the process handling a tool call
// runs in its own container, entirely separate from the filesystem of
// whatever MCP client is calling it (e.g. Claude's own sandbox). Writing
// attachment bytes to local disk here — as the original single-user stdio
// server did — would save the file somewhere the client can never read.
// Binary attachments must instead travel back as part of the MCP response
// itself, as a base64-encoded embedded resource.
const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024; // 20 MB decoded

/**
 * List attachments for a specific email
 * @param {object} args - Tool arguments
 * @param {string} args.messageId - The ID of the email message
 * @returns {object} - MCP response with attachment list
 */
async function handleListAttachments(args) {
  const messageId = args.messageId;

  if (!messageId) {
    return {
      content: [
        {
          type: 'text',
          text: 'Error: messageId is required',
        },
      ],
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Call Graph API to get attachments
    const endpoint = `/me/messages/${messageId}/attachments`;
    const params = {
      $select: 'id,name,contentType,size,isInline',
    };

    console.log(`Fetching attachments for message: ${messageId}`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      endpoint,
      null,
      params
    );

    if (!response.value || response.value.length === 0) {
      return {
        content: [
          {
            type: 'text',
            text: 'No attachments found for this email.',
          },
        ],
      };
    }

    // Format attachment list
    const attachmentList = response.value
      .map((att, index) => {
        const sizeKB = (att.size / 1024).toFixed(1);
        const inline = att.isInline ? ' [inline]' : '';
        return `${index + 1}. ${att.name}${inline}\n   Type: ${att.contentType}\n   Size: ${sizeKB} KB\n   ID: ${att.id}`;
      })
      .join('\n\n');

    return {
      content: [
        {
          type: 'text',
          text: `Found ${response.value.length} attachment(s):\n\n${attachmentList}`,
        },
      ],
    };
  } catch (error) {
    if (
      error.message === 'Authentication required' ||
      error.message === 'UNAUTHORIZED'
    ) {
      return {
        content: [
          {
            type: 'text',
            text: "Authentication required. Please use the 'authenticate' tool first.",
          },
        ],
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: `Error listing attachments: ${error.message}`,
        },
      ],
    };
  }
}

/**
 * Download a specific attachment — returns its content as an embedded MCP
 * resource (base64 blob) so the calling client can save or render it.
 * @param {object} args - Tool arguments
 * @param {string} args.messageId - The ID of the email message
 * @param {string} args.attachmentId - The ID of the attachment
 * @returns {object} - MCP response with the attachment as an embedded resource
 */
async function handleDownloadAttachment(args) {
  const { messageId, attachmentId } = args;

  if (!messageId || !attachmentId) {
    return {
      content: [
        {
          type: 'text',
          text: 'Error: Both messageId and attachmentId are required',
        },
      ],
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // First, get attachment metadata to get the filename and content
    const metadataEndpoint = `/me/messages/${messageId}/attachments/${attachmentId}`;
    console.log(`Fetching attachment metadata: ${attachmentId}`);

    const metadata = await callGraphAPI(
      accessToken,
      'GET',
      metadataEndpoint,
      null,
      {}
    );

    if (!metadata) {
      return {
        content: [
          {
            type: 'text',
            text: 'Error: Attachment not found',
          },
        ],
      };
    }

    const filename = metadata.name || 'attachment';
    const contentType = metadata.contentType || 'application/octet-stream';

    // For file attachments, the content is base64 encoded in contentBytes
    if (metadata['@odata.type'] === '#microsoft.graph.fileAttachment') {
      const contentBytes = metadata.contentBytes;

      if (!contentBytes) {
        return {
          content: [
            {
              type: 'text',
              text: 'Error: No content found in attachment',
            },
          ],
        };
      }

      const sizeBytes = Math.ceil((contentBytes.length * 3) / 4);
      if (sizeBytes > MAX_ATTACHMENT_BYTES) {
        const maxMB = (MAX_ATTACHMENT_BYTES / (1024 * 1024)).toFixed(0);
        return {
          content: [
            {
              type: 'text',
              text: `Attachment "${filename}" is too large to return (${(sizeBytes / (1024 * 1024)).toFixed(1)} MB, limit ${maxMB} MB). Download it directly from Outlook instead.`,
            },
          ],
        };
      }

      const sizeKB = (sizeBytes / 1024).toFixed(1);

      return {
        content: [
          {
            type: 'text',
            text: `Downloaded attachment:\n\nFilename: ${filename}\nType: ${contentType}\nSize: ${sizeKB} KB`,
          },
          {
            type: 'resource',
            resource: {
              uri: `attachment://${messageId}/${attachmentId}/${encodeURIComponent(filename)}`,
              mimeType: contentType,
              blob: contentBytes,
            },
          },
        ],
      };
    } else if (metadata['@odata.type'] === '#microsoft.graph.itemAttachment') {
      // Item attachments (embedded emails, calendar items) need different handling
      return {
        content: [
          {
            type: 'text',
            text: `This is an embedded item attachment (${metadata.name}). Item attachments cannot be downloaded as files directly. They contain embedded Outlook items like emails or calendar events.`,
          },
        ],
      };
    } else if (
      metadata['@odata.type'] === '#microsoft.graph.referenceAttachment'
    ) {
      // Reference attachments are links to cloud files
      return {
        content: [
          {
            type: 'text',
            text: `This is a reference attachment (cloud link):\n\nName: ${metadata.name}\nThis attachment is a link to a file stored in the cloud and cannot be downloaded directly.`,
          },
        ],
      };
    } else {
      return {
        content: [
          {
            type: 'text',
            text: `Unknown attachment type: ${metadata['@odata.type']}`,
          },
        ],
      };
    }
  } catch (error) {
    if (
      error.message === 'Authentication required' ||
      error.message === 'UNAUTHORIZED'
    ) {
      return {
        content: [
          {
            type: 'text',
            text: "Authentication required. Please use the 'authenticate' tool first.",
          },
        ],
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: `Error downloading attachment: ${error.message}`,
        },
      ],
    };
  }
}

/**
 * Get attachment content as base64 (useful for viewing in Claude)
 * @param {object} args - Tool arguments
 * @param {string} args.messageId - The ID of the email message
 * @param {string} args.attachmentId - The ID of the attachment
 * @returns {object} - MCP response with attachment content
 */
async function handleGetAttachmentContent(args) {
  const { messageId, attachmentId } = args;

  if (!messageId || !attachmentId) {
    return {
      content: [
        {
          type: 'text',
          text: 'Error: Both messageId and attachmentId are required',
        },
      ],
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `/me/messages/${messageId}/attachments/${attachmentId}`;
    console.log(`Fetching attachment content: ${attachmentId}`);

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, {});

    if (!response) {
      return {
        content: [
          {
            type: 'text',
            text: 'Error: Attachment not found',
          },
        ],
      };
    }

    const filename = response.name || 'attachment';
    const contentType = response.contentType || 'application/octet-stream';
    const sizeBytes = response.size || 0;
    const sizeKB = (sizeBytes / 1024).toFixed(1);

    // For file attachments, return metadata and indicate content is available
    if (response['@odata.type'] === '#microsoft.graph.fileAttachment') {
      // Check if it's a text-based file we can display
      const textTypes = [
        'text/',
        'application/json',
        'application/xml',
        'application/javascript',
      ];
      const isText = textTypes.some((t) => contentType.startsWith(t));

      if (isText && response.contentBytes) {
        const content = Buffer.from(response.contentBytes, 'base64').toString(
          'utf-8'
        );
        return {
          content: [
            {
              type: 'text',
              text: `Attachment: ${filename}\nType: ${contentType}\nSize: ${sizeKB} KB\n\n--- Content ---\n${content}`,
            },
          ],
        };
      }

      // For binary files, just return metadata
      return {
        content: [
          {
            type: 'text',
            text: `Attachment: ${filename}\nType: ${contentType}\nSize: ${sizeKB} KB\n\nThis is a binary file. Use action='download' to retrieve its content.`,
          },
        ],
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: `Attachment: ${filename}\nType: ${response['@odata.type']}\nSize: ${sizeKB} KB`,
        },
      ],
    };
  } catch (error) {
    if (
      error.message === 'Authentication required' ||
      error.message === 'UNAUTHORIZED'
    ) {
      return {
        content: [
          {
            type: 'text',
            text: "Authentication required. Please use the 'authenticate' tool first.",
          },
        ],
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: `Error getting attachment content: ${error.message}`,
        },
      ],
    };
  }
}

module.exports = {
  handleListAttachments,
  handleDownloadAttachment,
  handleGetAttachmentContent,
};
