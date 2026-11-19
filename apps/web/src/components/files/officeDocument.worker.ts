import { msg } from "gt-react";
import { m } from "~/i18n/messages";
import {
  parseOfficeDocument,
  type OfficeWorkerRequest,
  type OfficeWorkerResponse,
} from "./officeDocumentParser";

const sendResponse: (response: OfficeWorkerResponse) => void = self.postMessage.bind(self);

self.addEventListener("message", (event: MessageEvent<OfficeWorkerRequest>) => {
  const { id, format, buffer } = event.data;
  let response: OfficeWorkerResponse;
  try {
    if (format !== "xlsx" && format !== "xls" && format !== "pptx") {
      throw new Error("Preview supports .xlsx, binary .xls, and basic .pptx slides.");
    }
    response = {
      id,
      ok: true,
      document: parseOfficeDocument(new Uint8Array(buffer), format),
    };
  } catch (error) {
    response = {
      id,
      ok: false,
      error: error instanceof Error ? error.message : m(msg("Could not read this Office document.")),
    };
  }
  sendResponse(response);
});

sendResponse({ type: "ready" });
