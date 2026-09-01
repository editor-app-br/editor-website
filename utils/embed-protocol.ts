/**
 * Protocol definition for Thin-Embed integration between Host and editor.app.br
 */

export type EmbedDocumentType = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'doc' | 'xls' | 'ppt' | 'txt' | 'csv';

export type HostToEditorMessage =
  | {
      type: 'open';
      requestId: string;
      fileName: string;
      fileType: EmbedDocumentType;
      editing: boolean;
      theme?: 'dark' | 'light' | 'system';
      lang?: string;
      bytes: ArrayBuffer;
    }
  | {
      type: 'save';
      requestId: string;
    }
  | {
      type: 'setTheme';
      theme: 'dark' | 'light';
    }
  | {
      type: 'print';
    }
  | {
      type: 'destroy';
    };

export type EditorToHostMessage =
  | {
      type: 'ready';
      version: string;
    }
  | {
      type: 'dirty';
      value: boolean;
    }
  | {
      type: 'saved';
      requestId?: string;
      fileName: string;
      fileType: string;
      mime: string;
      bytes: ArrayBuffer;
    }
  | {
      type: 'error';
      message: string;
    }
  | {
      type: 'escape';
    };
