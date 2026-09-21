import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { act, render, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { Suspense } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";

import KBDetailPage from "./page";
import { apiClient } from "@/lib/api-client";
import { DEFAULT_INGESTION_CONFIG } from "@/lib/ingestion-config";
import { useOrgStore } from "@/stores";
import { pick } from "@/test-utils/file-picker";
import type { KnowledgeBase } from "@/types";

vi.mock("@/lib/api-client", () => ({
  apiClient: { get: vi.fn(), post: vi.fn(), patch: vi.fn(), delete: vi.fn(), raw: vi.fn() },
  ApiError: class ApiError extends Error {},
}));
vi.mock("sonner", () => ({ toast: { success: vi.fn(), error: vi.fn() } }));
// Key-returning translator, cached per namespace. The cache is load-bearing rather
// than tidy: `useKBDetail` puts `t` in `refresh`'s dependencies and the page runs
// `useEffect(() => refresh(), [refresh])`, so a translator rebuilt per call re-fires
// that effect forever and every test in the file times out (#446). `rich`/`markup`/
// `has` hang off the same function, because a component under this tree reads a
// message with a tag in it (#612).
vi.mock("next-intl", async () => {
  const build = (await import("@/test-utils/intl")).keyTranslations((ns, key) => `${ns}.${key}`);
  const cache = new Map<string, ReturnType<typeof build>>();
  return {
    useLocale: () => "en",
    useTranslations: (ns: string) => {
      let translate = cache.get(ns);
      if (translate === undefined) {
        translate = build(ns);
        cache.set(ns, translate);
      }
      return translate;
    },
  };
});

// `collections:edit`, not the `collections:view` the sibling specs hold:
// `handleFiles` returns early without it, so a Viewer's page would satisfy every
// assertion here by never uploading anything at all.
const perms = new Set<string>(["collections:view", "collections:edit"]);
vi.mock("@/hooks/use-permissions", () => ({
  usePermissions: () => ({ can: (p: string) => perms.has(p) }),
}));

const ORG_ID = "org-1";
const KB_ID = "5eacffcc-873e-42fe-a73a-32cd19322d00";

const KB: KnowledgeBase = {
  id: KB_ID,
  name: "Handbook",
  description: null,
  collection_name: "handbook",
  scope: "org",
  organization_id: ORG_ID,
  owner_user_id: null,
  is_default: false,
  ingestion_config: DEFAULT_INGESTION_CONFIG,
  embedding_model: "text-embedding-3-large",
  embedding_provider: "openrouter",
  embedding_secret_id: null,
  embedding_endpoint_id: null,
  embedding_dim: 3072,
  created_at: "2026-07-01T00:00:00Z",
  updated_at: null,
  document_count: 0,
  indexed_count: 0,
  chunk_count: 0,
};

/**
 * A fake `XMLHttpRequest`.
 *
 * The upload goes around `apiClient` deliberately, to read byte-level progress
 * off an XHR, so this is where the request shows up rather than on the mocked
 * client. Same shape as `use-knowledge-bases.test.tsx`'s.
 */
interface FakeXhr {
  open: ReturnType<typeof vi.fn>;
  setRequestHeader: ReturnType<typeof vi.fn>;
  send: ReturnType<typeof vi.fn>;
  withCredentials: boolean;
  status: number;
  responseText: string;
  upload: { onprogress?: (event: ProgressEvent) => void; onload?: () => void };
  onload?: () => void;
  onerror?: () => void;
}

/** One entry per upload started, in order. */
let xhrs: FakeXhr[];

function stubXhr() {
  xhrs = [];
  // A class, not `vi.fn(() => ({...}))`: the code under test calls
  // `new XMLHttpRequest()`, which an arrow function cannot answer.
  class FakeXhrImpl implements FakeXhr {
    open = vi.fn();
    send = vi.fn();
    setRequestHeader = vi.fn();
    withCredentials = false;
    status = 201;
    responseText = "{}";
    upload: FakeXhr["upload"] = {};

    constructor() {
      xhrs.push(this);
    }
  }
  vi.stubGlobal("XMLHttpRequest", FakeXhrImpl);
}

function mockApi() {
  vi.mocked(apiClient.get).mockImplementation((endpoint: string) => {
    if (endpoint === `/kb/${KB_ID}`) return Promise.resolve(KB);
    return Promise.resolve({ items: [], total: 0 });
  });
}

/**
 * A pre-fulfilled thenable, so `use(params)` reads it synchronously instead of
 * suspending - the same fast path React takes once a promise has settled.
 */
function fulfilled<T>(value: T): Promise<T> {
  const thenable = Promise.resolve(value) as Promise<T> & { status: string; value: T };
  thenable.status = "fulfilled";
  thenable.value = value;
  return thenable;
}

async function renderPage() {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  const view = render(
    <QueryClientProvider client={client}>
      <Suspense fallback={null}>
        <KBDetailPage params={fulfilled({ id: KB_ID })} />
      </Suspense>
    </QueryClientProvider>,
  );
  await act(async () => {});
  return view;
}

/**
 * The page's one hidden picker.
 *
 * Awaited rather than read straight off the container: the page draws a skeleton
 * until its collection arrives, so a synchronous `querySelector` races the first
 * render and hands back `null`.
 */
async function picker(container: HTMLElement): Promise<HTMLInputElement> {
  let input: HTMLInputElement | null = null;
  await waitFor(() => {
    input = container.querySelector<HTMLInputElement>('input[type="file"]');
    expect(input).not.toBeNull();
  });
  return input!;
}

beforeEach(() => {
  // The page writes its tab into the URL, and jsdom's location persists across
  // tests in a file. A browser gets a fresh URL per navigation.
  window.history.replaceState({}, "", "/");
  vi.clearAllMocks();
  vi.unstubAllGlobals();
  stubXhr();
  useOrgStore.setState({ activeOrgId: ORG_ID });
  mockApi();
});

describe("the knowledge base's file picker", () => {
  it("clears the input after a pick, so the same document can be picked again", async () => {
    // The defect this file exists for: nothing reset the input, and a file input
    // fires `change` only when the selection *changes*. Picking the same document
    // a second time therefore fired nothing at all - no request, no toast, no
    // progress row - which is exactly what somebody does after an upload that
    // failed for a reason they have since fixed.
    //
    // Asserted on the value rather than on a second upload on purpose: an empty
    // value is the browser's precondition for that second `change`, and jsdom
    // fires `change` either way, so the event a browser withholds cannot be
    // observed from here. This is that precondition, stated directly.
    const { container } = await renderPage();
    const input = await picker(container);

    await userEvent.upload(input, new File(["hi"], "notes.txt", { type: "text/plain" }));

    await waitFor(() => expect(xhrs).toHaveLength(1));
    expect(input.value).toBe("");
  });

  it("copies the picked files before clearing, so the upload still sees them", async () => {
    // The other half, and the one the reset above is a single line away from
    // breaking: `e.target.files` is the input's own live `FileList`, so clearing
    // the input empties that same object in place in Blink and WebKit. `pick`
    // models that clear; `userEvent.upload` cannot, which is how the chat
    // composer's picker shipped dead with a green suite.
    const { container } = await renderPage();

    pick(await picker(container), new File(["hi"], "notes.txt", { type: "text/plain" }));

    await waitFor(() => expect(xhrs[0]?.send).toHaveBeenCalled());
    expect(xhrs[0]!.open).toHaveBeenCalledWith("POST", `/api/kb/${KB_ID}/documents`);
    const body = xhrs[0]!.send.mock.calls[0]![0] as FormData;
    expect((body.get("file") as File).name).toBe("notes.txt");
  });

  it("clears the input on an empty pick too, without uploading", async () => {
    // The reset sits ahead of the empty check rather than after it, so a change
    // that selected nothing still leaves the input able to fire again - the pick
    // it would otherwise refuse is the one right after somebody changed their
    // mind in the dialog.
    const { container } = await renderPage();

    const picked = pick(await picker(container));

    await act(async () => {});
    expect(picked.resets).toBe(1);
    expect(xhrs).toEqual([]);
  });
});
