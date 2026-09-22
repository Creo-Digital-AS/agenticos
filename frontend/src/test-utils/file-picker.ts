import { fireEvent } from "@testing-library/react";

/**
 * A `FileList` that empties *in place* when the input is reset, as Blink's does.
 *
 * This is the one browser behaviour the suite cannot get from jsdom, and the
 * reason the chat composer's picker shipped broken with every test around it
 * green: jsdom hands the input a fresh list on reset, so a handler that captures
 * `input.files` and *then* sets `input.value = ""` still sees a populated list
 * under test and an empty one in Chrome, Edge and Safari. Modelled rather than
 * trusted.
 */
function blinkFileList(...files: File[]): FileList {
  const list = {
    length: files.length,
    item: (index: number): File | null => files[index] ?? null,
    clear(): void {
      this.length = 0;
    },
    *[Symbol.iterator](): Generator<File> {
      for (let index = 0; index < this.length; index++) yield files[index]!;
    },
  };
  files.forEach((file, index) => Object.defineProperty(list, index, { value: file }));
  // A test double for a host object with no constructor reachable from script.
  return list as unknown as FileList;
}

/** What a pick did, once the handler it fired has run. */
export interface Picked {
  /**
   * How many times the handler cleared the input.
   *
   * A file input fires `change` only when the selection *changes*, so an
   * unreset input silently refuses the same file twice. That clear is the
   * behaviour under test wherever a picker can be given the same file again,
   * and a count is the only way to see it: the `value` getter below answers
   * `""` either way, exactly as a browser's does once nothing is selected.
   */
  resets: number;
}

/**
 * Pick `files` through `input`, with `input.value = ""` clearing `input.files`
 * in place the way a browser does.
 *
 * `userEvent.upload` cannot model that clear, so a handler whose copy and reset
 * are in the wrong order passes under it and drops every file in a real
 * browser. Reach for this wherever that ordering is the thing under test;
 * `userEvent.upload` is still the right call for everything else.
 *
 * Synchronous: `change` is dispatched before this returns, so the count in the
 * result is final even when the handler goes on to await an upload.
 */
export function pick(input: HTMLInputElement, ...files: File[]): Picked {
  const list = blinkFileList(...files);
  const picked: Picked = { resets: 0 };
  Object.defineProperty(input, "files", { configurable: true, get: () => list });
  Object.defineProperty(input, "value", {
    configurable: true,
    get: () => "",
    set: () => {
      picked.resets += 1;
      (list as unknown as { clear: () => void }).clear();
    },
  });
  fireEvent.change(input);
  return picked;
}
