/**
 * Word.run wrappers — all document mutations live here.
 *
 * Insertion strategy: tags are inserted as plain text in a SINGLE insertText
 * call, so each lands in one <w:r> run (the engine's run-coalescer is the
 * fallback for hand-edits, not the primary path — see plan "Learnings").
 */

import { ConditionalSpec, conditionalTags, inverseTags, loopRowCellTexts } from "./tags";

/** Insert a scalar/built-in tag at the cursor (or replacing the selection). */
export async function insertScalar(tagText: string): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(tagText, Word.InsertLocation.replace);
    selection.getRange(Word.RangeLocation.end).select();
    await context.sync();
  });
}

/**
 * Insert a row-scope loop table per Phase H §3.7.1:
 * header row with field labels + one data row where `{{#Rel}}` opens in the
 * first cell and `{{/Rel}}` closes in the last cell of the SAME row.
 */
export async function insertLoopTable(
  relationshipName: string,
  columns: { inLoopKey: string; label: string }[],
): Promise<void> {
  const headers = columns.map((c) => c.label);
  const dataRow = loopRowCellTexts(
    relationshipName,
    columns.map((c) => c.inLoopKey),
  );
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const table = selection.insertTable(2, columns.length, Word.InsertLocation.after, [
      headers,
      dataRow,
    ]);
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
    table.getRange(Word.RangeLocation.after).select();
    await context.sync();
  });
}

/** Insert an inline conditional around the current selection (or placeholder text). */
export async function insertConditional(spec: ConditionalSpec): Promise<void> {
  const { open, elseTag, close } = conditionalTags(spec);
  await wrapSelection(open, elseTag ? `${elseTag}otherwise…${close}` : close);
}

/** Insert an inverse (show-when-blank) block around the selection. */
export async function insertInverse(fieldKey: string): Promise<void> {
  const { open, close } = inverseTags(fieldKey);
  await wrapSelection(open, close);
}

/** Wrap the selection with already-built conditional tags (e.g. compound if). */
export async function insertConditionalTags(
  open: string,
  elseTag: string | undefined,
  close: string,
): Promise<void> {
  await wrapSelection(open, elseTag ? `${elseTag}otherwise…${close}` : close);
}

/** Insert an image merge tag at the cursor. */
export async function insertImage(tagText: string): Promise<void> {
  await insertScalar(tagText);
}

/**
 * Wraps the current selection with open/close tag text. With a collapsed
 * cursor, inserts `open` + placeholder + `close` and selects the placeholder
 * so the user types content immediately.
 */
async function wrapSelection(openText: string, closeText: string): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("isEmpty");
    await context.sync();

    if (selection.isEmpty) {
      const inserted = selection.insertText(
        `${openText}content${closeText}`,
        Word.InsertLocation.replace,
      );
      // Select the placeholder word "content" for immediate overtype.
      const search = inserted.search("content", { matchCase: true });
      search.load("items");
      await context.sync();
      search.items[0]?.select();
      await context.sync();
    } else {
      selection.insertText(openText, Word.InsertLocation.start);
      selection.insertText(closeText, Word.InsertLocation.end);
      selection.getRange(Word.RangeLocation.end).select();
      await context.sync();
    }
  });
}

/** Full text of the body — used by mock-mode lint to scan for tags. */
export async function getDocumentText(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text;
  });
}

/**
 * The current document as base64 (compressed .docx), via
 * Office.context.document.getFileAsync — step 1 payload of the two-step save.
 */
export function getDocumentAsBase64(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4 * 1024 * 1024 },
      (fileResult) => {
        if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(fileResult.error?.message ?? "getFileAsync failed"));
          return;
        }
        const file = fileResult.value;
        const slices: number[][] = [];
        let received = 0;

        const readSlice = (index: number): void => {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error(sliceResult.error?.message ?? "getSliceAsync failed"));
              return;
            }
            slices[index] = sliceResult.value.data as number[];
            received += 1;
            if (received === file.sliceCount) {
              file.closeAsync();
              resolve(bytesToBase64(slices.flat()));
            } else {
              readSlice(index + 1);
            }
          });
        };
        readSlice(0);
      },
    );
  });
}

function bytesToBase64(bytes: number[]): string {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.slice(i, i + chunk));
  }
  return btoa(binary);
}

/** Suggested template name from the doc title, e.g. "Quote Template.docx" → "Quote Template". */
export function suggestedTemplateName(): string {
  const url = Office.context.document.url ?? "";
  const last = url.split(/[\\/]/).pop() ?? "";
  const stripped = last.replace(/\.(docx|docm)$/i, "");
  return stripped || "Untitled Template";
}
