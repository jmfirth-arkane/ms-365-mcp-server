import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'get-mail-attachment',
        method: 'GET',
        path: '/me/messages/:message-id/attachments/:attachment-id',
        description: 'Get a mail attachment',
        parameters: [
          { name: 'message-id', type: 'Path' },
          { name: 'attachment-id', type: 'Path' },
        ],
      },
      {
        alias: 'list-mail-attachments',
        method: 'GET',
        path: '/me/messages/:message-id/attachments',
        description: 'List mail attachments',
        parameters: [
          { name: 'message-id', type: 'Path' },
        ],
      },
      {
        alias: 'get-onenote-page-content',
        method: 'GET',
        path: '/me/onenote/pages/:onenotePage-id/content',
        description: 'Get OneNote page content',
        parameters: [
          { name: 'onenotePage-id', type: 'Path' },
        ],
      },
    ],
  },
}));

describe('Resource parts gating', () => {
  let graphClient: GraphClient;
  let mockGraphRequest: ReturnType<typeof vi.fn>;

  // Capture the tool handlers registered by registerGraphTools
  function registerAndCapture(resourceParts: boolean): Map<string, Function> {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = new Map<string, Function>();

    vi.spyOn(server, 'tool').mockImplementation((...args: any[]) => {
      // server.tool(name, description, schema, annotations, handler)
      const name = args[0] as string;
      const handler = args[args.length - 1] as Function;
      handlers.set(name, handler);
    });

    registerGraphTools(server, graphClient, false, undefined, false, resourceParts);
    return handlers;
  }

  beforeEach(() => {
    mockGraphRequest = vi.fn();
    graphClient = { graphRequest: mockGraphRequest } as unknown as GraphClient;
  });

  describe('get-mail-attachment', () => {
    const singleAttachmentResponse = {
      content: [
        {
          text: JSON.stringify({
            id: 'att-1',
            name: 'report.pdf',
            contentType: 'application/pdf',
            size: 1024,
            contentBytes: 'JVBERi0xLjQK',
          }),
        },
      ],
    };

    it('resourceParts=false: returns text-only with contentBytes in JSON', async () => {
      mockGraphRequest.mockResolvedValue(singleAttachmentResponse);
      const handlers = registerAndCapture(false);
      const handler = handlers.get('get-mail-attachment')!;

      const result = await handler({ 'message-id': 'msg-1', 'attachment-id': 'att-1' });

      expect(result.content).toHaveLength(1);
      expect(result.content[0].type).toBe('text');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.contentBytes).toBe('JVBERi0xLjQK');
    });

    it('resourceParts=true: returns text metadata + resource blob', async () => {
      mockGraphRequest.mockResolvedValue(singleAttachmentResponse);
      const handlers = registerAndCapture(true);
      const handler = handlers.get('get-mail-attachment')!;

      const result = await handler({ 'message-id': 'msg-1', 'attachment-id': 'att-1' });

      expect(result.content).toHaveLength(2);

      // Text metadata (no contentBytes)
      expect(result.content[0].type).toBe('text');
      const metadata = JSON.parse(result.content[0].text);
      expect(metadata.contentBytes).toBeUndefined();
      expect(metadata.name).toBe('report.pdf');

      // Resource blob
      expect(result.content[1].type).toBe('resource');
      expect(result.content[1].resource.blob).toBe('JVBERi0xLjQK');
      expect(result.content[1].resource.mimeType).toBe('application/pdf');
      expect(result.content[1].resource.uri).toContain('msgraph://mail/messages/msg-1/attachments/att-1/');
    });
  });

  describe('list-mail-attachments', () => {
    const listAttachmentsResponse = {
      content: [
        {
          text: JSON.stringify({
            value: [
              {
                id: 'att-1',
                name: 'file1.pdf',
                contentType: 'application/pdf',
                contentBytes: 'base64pdf',
              },
              {
                id: 'att-2',
                name: 'file2.png',
                contentType: 'image/png',
                contentBytes: 'base64png',
              },
            ],
          }),
        },
      ],
    };

    it('resourceParts=false: returns text-only JSON with contentBytes in value array', async () => {
      mockGraphRequest.mockResolvedValue(listAttachmentsResponse);
      const handlers = registerAndCapture(false);
      const handler = handlers.get('list-mail-attachments')!;

      const result = await handler({ 'message-id': 'msg-1' });

      expect(result.content).toHaveLength(1);
      expect(result.content[0].type).toBe('text');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.value).toHaveLength(2);
      expect(parsed.value[0].contentBytes).toBe('base64pdf');
      expect(parsed.value[1].contentBytes).toBe('base64png');
    });

    it('resourceParts=true: returns interleaved text + resource blob parts', async () => {
      mockGraphRequest.mockResolvedValue(listAttachmentsResponse);
      const handlers = registerAndCapture(true);
      const handler = handlers.get('list-mail-attachments')!;

      const result = await handler({ 'message-id': 'msg-1' });

      // 2 attachments with contentBytes = 2 text + 2 resource = 4 parts
      expect(result.content).toHaveLength(4);

      expect(result.content[0].type).toBe('text');
      expect(result.content[1].type).toBe('resource');
      expect(result.content[1].resource.blob).toBe('base64pdf');

      expect(result.content[2].type).toBe('text');
      expect(result.content[3].type).toBe('resource');
      expect(result.content[3].resource.blob).toBe('base64png');
    });
  });

  describe('get-onenote-page-content', () => {
    const html = '<html><body><h1>My Page</h1></body></html>';
    const onenoteResponse = {
      content: [
        {
          text: JSON.stringify({ message: 'OK!', rawResponse: html }),
        },
      ],
    };

    it('resourceParts=false: returns text-only JSON (the rawResponse wrapper)', async () => {
      mockGraphRequest.mockResolvedValue(onenoteResponse);
      const handlers = registerAndCapture(false);
      const handler = handlers.get('get-onenote-page-content')!;

      const result = await handler({ 'onenotePage-id': 'page-1' });

      expect(result.content).toHaveLength(1);
      expect(result.content[0].type).toBe('text');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.rawResponse).toBe(html);
    });

    it('resourceParts=true: returns ResourceTextContent with text/html', async () => {
      mockGraphRequest.mockResolvedValue(onenoteResponse);
      const handlers = registerAndCapture(true);
      const handler = handlers.get('get-onenote-page-content')!;

      const result = await handler({ 'onenotePage-id': 'page-1' });

      expect(result.content).toHaveLength(1);
      expect(result.content[0].type).toBe('resource');
      expect(result.content[0].resource.mimeType).toBe('text/html');
      expect(result.content[0].resource.text).toBe(html);
      expect(result.content[0].resource.uri).toBe('msgraph://onenote/pages/page-1/content');
    });
  });
});
