import { describe, expect, it, vi } from 'vitest';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
  },
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [],
  },
}));

import { extractAttachmentResources } from '../src/graph-tools.js';

describe('extractAttachmentResources', () => {
  const baseParams = { 'message-id': 'msg-123' };

  describe('single attachment (get-mail-attachment)', () => {
    it('should return text metadata + resource blob for attachment with contentBytes', () => {
      const data = {
        id: 'att-1',
        name: 'report.pdf',
        contentType: 'application/pdf',
        size: 204800,
        isInline: false,
        contentBytes: 'JVBERi0xLjQK',
      };

      const result = extractAttachmentResources(data, baseParams);

      expect(result).toHaveLength(2);

      // First part: text metadata without contentBytes
      expect(result[0]).toEqual({
        type: 'text',
        text: JSON.stringify(
          { id: 'att-1', name: 'report.pdf', contentType: 'application/pdf', size: 204800, isInline: false },
          null,
          2
        ),
      });

      // Second part: resource blob
      expect(result[1]).toEqual({
        type: 'resource',
        resource: {
          uri: 'msgraph://mail/messages/msg-123/attachments/att-1/report.pdf',
          mimeType: 'application/pdf',
          blob: 'JVBERi0xLjQK',
        },
      });
    });

    it('should return text-only for attachment without contentBytes (reference attachment)', () => {
      const data = {
        id: 'att-2',
        name: 'shared-doc.docx',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        size: 0,
        isInline: false,
        sourceUrl: 'https://sharepoint.example.com/doc',
      };

      const result = extractAttachmentResources(data, baseParams);

      expect(result).toHaveLength(1);
      expect(result[0].type).toBe('text');
      const parsed = JSON.parse((result[0] as any).text);
      expect(parsed.id).toBe('att-2');
      expect(parsed.name).toBe('shared-doc.docx');
      expect(parsed.sourceUrl).toBe('https://sharepoint.example.com/doc');
      expect(parsed.contentBytes).toBeUndefined();
    });

    it('should URL-encode filename in URI', () => {
      const data = {
        id: 'att-3',
        name: 'file with spaces & (special).pdf',
        contentType: 'application/pdf',
        contentBytes: 'base64data',
      };

      const result = extractAttachmentResources(data, baseParams);
      const resource = result[1] as any;
      expect(resource.resource.uri).toContain(encodeURIComponent('file with spaces & (special).pdf'));
    });

    it('should use default mimeType when contentType is missing', () => {
      const data = {
        id: 'att-4',
        name: 'mystery-file',
        contentBytes: 'base64data',
      };

      const result = extractAttachmentResources(data, baseParams);
      const resource = result[1] as any;
      expect(resource.resource.mimeType).toBe('application/octet-stream');
    });

    it('should use "unknown" for missing message-id', () => {
      const data = {
        id: 'att-5',
        name: 'file.txt',
        contentType: 'text/plain',
        contentBytes: 'aGVsbG8=',
      };

      const result = extractAttachmentResources(data, {});
      const resource = result[1] as any;
      expect(resource.resource.uri).toContain('msgraph://mail/messages/unknown/attachments/');
    });
  });

  describe('list of attachments (list-mail-attachments)', () => {
    it('should return interleaved text + resource pairs for multiple attachments', () => {
      const data = {
        value: [
          {
            id: 'att-1',
            name: 'report.pdf',
            contentType: 'application/pdf',
            size: 204800,
            contentBytes: 'JVBERi0xLjQK',
          },
          {
            id: 'att-2',
            name: 'image.png',
            contentType: 'image/png',
            size: 51200,
            contentBytes: 'iVBORw0KGgo=',
          },
        ],
      };

      const result = extractAttachmentResources(data, baseParams);

      expect(result).toHaveLength(4);

      // First attachment: text + resource
      expect(result[0].type).toBe('text');
      expect(JSON.parse((result[0] as any).text).name).toBe('report.pdf');
      expect(JSON.parse((result[0] as any).text).contentBytes).toBeUndefined();

      expect(result[1]).toEqual({
        type: 'resource',
        resource: {
          uri: 'msgraph://mail/messages/msg-123/attachments/att-1/report.pdf',
          mimeType: 'application/pdf',
          blob: 'JVBERi0xLjQK',
        },
      });

      // Second attachment: text + resource
      expect(result[2].type).toBe('text');
      expect(JSON.parse((result[2] as any).text).name).toBe('image.png');

      expect(result[3]).toEqual({
        type: 'resource',
        resource: {
          uri: 'msgraph://mail/messages/msg-123/attachments/att-2/image.png',
          mimeType: 'image/png',
          blob: 'iVBORw0KGgo=',
        },
      });
    });

    it('should fallback to text JSON for empty value array', () => {
      const data = { value: [] };

      const result = extractAttachmentResources(data, baseParams);

      expect(result).toHaveLength(1);
      expect(result[0].type).toBe('text');
      expect((result[0] as any).text).toBe(JSON.stringify(data, null, 2));
    });

    it('should handle mixed list (some with contentBytes, some without)', () => {
      const data = {
        value: [
          {
            id: 'att-1',
            name: 'file.pdf',
            contentType: 'application/pdf',
            contentBytes: 'base64data',
          },
          {
            id: 'att-2',
            name: 'ref-doc.docx',
            contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            sourceUrl: 'https://sharepoint.example.com/doc',
          },
        ],
      };

      const result = extractAttachmentResources(data, baseParams);

      // att-1: text + resource = 2 parts
      // att-2: text only = 1 part
      expect(result).toHaveLength(3);

      expect(result[0].type).toBe('text');
      expect(result[1].type).toBe('resource');
      expect(result[2].type).toBe('text');

      expect(JSON.parse((result[2] as any).text).name).toBe('ref-doc.docx');
      expect(JSON.parse((result[2] as any).text).sourceUrl).toBe('https://sharepoint.example.com/doc');
    });
  });

  describe('fallback behavior', () => {
    it('should fallback to text for unrecognized data structure', () => {
      const data = { someField: 'value', anotherField: 123 };

      const result = extractAttachmentResources(data, baseParams);

      expect(result).toHaveLength(1);
      expect(result[0].type).toBe('text');
      expect((result[0] as any).text).toBe(JSON.stringify(data, null, 2));
    });

    it('should fallback for data with id but no name', () => {
      const data = { id: 'something', otherField: true };

      const result = extractAttachmentResources(data, baseParams);

      expect(result).toHaveLength(1);
      expect(result[0].type).toBe('text');
    });
  });
});
