import { T, Var, useGT } from "gt-react";
import { useState, useSyncExternalStore } from "react";

import { Button } from "~/components/ui/button";
import { FileSurfaceFailure, FileSurfaceLoading, FileSurfaceNotice } from "./fileSurfaceChrome";
import { createOfficeDocumentClient, type OfficeDocumentSource } from "./officeDocumentClient";
import type { OfficeDocument } from "./officeDocumentParser";

function columnLabel(index: number): string {
  return index < 26 ? String.fromCharCode(65 + index) : `A${String.fromCharCode(65 + index - 26)}`;
}

function DocumentContents({ document }: { readonly document: OfficeDocument }) {
  const gt = useGT();
  const [selected, setSelected] = useState(0);
  const [outline, setOutline] = useState(false);
  if (document.kind === "presentation") {
    const slide = document.slides[selected];
    const paragraphs = slide?.text.map((text, paragraph) => ({ text, paragraph }));
    const layers = slide?.elements.map((element, layer) => ({ element, layer }));
    return (
      <>
        <nav
          aria-label={gt("Slides")}
          className="flex min-w-0 shrink-0 flex-nowrap gap-1 overflow-x-auto border-b p-2"
        >
          {document.slides.map((entry, index) => (
            <Button
              key={entry.number}
              type="button"
              variant="ghost"
              size="sm"
              className="shrink-0"
              aria-pressed={selected === index}
              onClick={() => setSelected(index)}
            >
              <T>
                Slide <Var>{entry.number}</Var>
              </T>
            </Button>
          ))}
        </nav>
        <div className="flex shrink-0 gap-1 border-b p-2" aria-label={gt("Presentation view")}>
          <Button
            type="button"
            variant="ghost"
            size="sm"
            aria-pressed={!outline}
            onClick={() => setOutline(false)}
          >
            <T>
              <T>Slide preview</T>
            </T>
          </Button>
          <Button
            type="button"
            variant="ghost"
            size="sm"
            aria-pressed={outline}
            onClick={() => setOutline(true)}
          >
            <T>
              <T>Text outline</T>
            </T>
          </Button>
        </div>
        {document.truncated || slide?.truncated ? (
          <FileSurfaceNotice>
            <T>
              <T>Some slide text was omitted to keep the preview small.</T>
            </T>
          </FileSurfaceNotice>
        ) : null}
        <section
          aria-label={
            outline
              ? gt("Slide {number} text", { number: slide?.number ?? 1 })
              : gt("Slide {number} preview", { number: slide?.number ?? 1 })
          }
          className="scrollbar-gutter-both min-h-0 flex-1 overflow-auto p-4"
        >
          {outline ? (
            paragraphs?.length ? (
              paragraphs.map(({ text, paragraph }) => (
                <p key={paragraph} className="mb-3 whitespace-pre-wrap break-words">
                  {text}
                </p>
              ))
            ) : (
              <p className="text-muted-foreground">
                <T>No supported text on this slide.</T>
              </p>
            )
          ) : slide ? (
            <div
              className="@container relative w-full overflow-hidden border text-black shadow-sm"
              style={{
                aspectRatio: `${slide.width} / ${slide.height}`,
                backgroundColor: slide.background ?? "#ffffff",
              }}
            >
              {layers?.map(({ element, layer }) => {
                const geometry = {
                  left: `${(100 * element.x) / slide.width}%`,
                  top: `${(100 * element.y) / slide.height}%`,
                  width: `${(100 * element.width) / slide.width}%`,
                  height: `${(100 * element.height) / slide.height}%`,
                  transform: element.rotation ? `rotate(${element.rotation}deg)` : undefined,
                };
                if (element.kind === "text")
                  return (
                    <div
                      key={layer}
                      className="absolute overflow-hidden whitespace-pre-wrap break-words"
                      style={{
                        ...geometry,
                        fontSize: `${((element.fontSize * 12700) / slide.width) * 100}cqw`,
                        color: element.color,
                        fontWeight: element.bold ? "bold" : "normal",
                        textAlign: element.align,
                        lineHeight: 1.2,
                      }}
                    >
                      {element.text}
                    </div>
                  );
                if (element.kind === "image")
                  return /^data:image\/(?:png|jpeg|gif|webp);base64,/.test(element.src) ? (
                    <img
                      key={layer}
                      alt={element.alt}
                      src={element.src}
                      className="absolute object-contain"
                      style={geometry}
                    />
                  ) : null;
                return (
                  <div
                    key={layer}
                    aria-hidden="true"
                    className="absolute"
                    style={{
                      ...geometry,
                      borderRadius: element.shape === "ellipse" ? "50%" : undefined,
                      backgroundColor: element.fill,
                      border: element.stroke ? `1px solid ${element.stroke}` : undefined,
                    }}
                  />
                );
              })}
            </div>
          ) : null}
        </section>
      </>
    );
  }
  const sheet = document.sheets[selected];
  const columns = Math.min(30, sheet?.totalColumns ?? 0);
  const rows = sheet?.rows.slice(0, 100).map((cells, row) => ({ cells, rowNumber: row + 1 }));
  return (
    <>
      <nav
        aria-label={gt("Worksheets")}
        className="flex min-w-0 shrink-0 flex-nowrap gap-1 overflow-x-auto border-b p-2"
      >
        {document.sheets.map((entry, index) => (
          <Button
            key={entry.name}
            type="button"
            variant="ghost"
            size="sm"
            className="shrink-0"
            aria-pressed={selected === index}
            onClick={() => setSelected(index)}
          >
            <span className="max-w-48 truncate">{entry.name}</span>
          </Button>
        ))}
      </nav>
      {sheet?.truncated ? (
        <FileSurfaceNotice>
          <T>
            <T>Showing at most 100 rows, 30 columns, and 2,000 characters per cell.</T>
          </T>
        </FileSurfaceNotice>
      ) : null}
      <div className="scrollbar-gutter-both min-h-0 flex-1 overflow-auto">
        {sheet?.rows.length ? (
          <table
            aria-label={gt("{name} cell values", { name: sheet.name })}
            className="w-full border-collapse text-xs"
          >
            <thead>
              <tr>
                <th scope="col" className="border p-2">
                  <T>
                    <T>Row</T>
                  </T>
                </th>
                {Array.from({ length: columns }, (_, index) => (
                  <th key={columnLabel(index)} scope="col" className="border p-2 text-left">
                    {columnLabel(index)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows?.map(({ cells, rowNumber }) => (
                <tr key={rowNumber}>
                  <th
                    scope="row"
                    className="border p-2 text-right font-normal text-muted-foreground"
                  >
                    {rowNumber}
                  </th>
                  {Array.from({ length: columns }, (_, column) => (
                    <td
                      key={columnLabel(column)}
                      className="max-w-80 border p-2 align-top whitespace-pre-wrap break-words"
                    >
                      {cells[column]?.slice(0, 2000) ?? ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <p className="p-4 text-muted-foreground">
            <T>This worksheet is empty.</T>
          </p>
        )}
      </div>
    </>
  );
}

export default function OfficeDocumentPreview(
  props: OfficeDocumentSource & {
    readonly onRetry?: (() => void) | undefined;
  },
) {
  const [revision, setRevision] = useState(0);
  const source = { name: props.name, format: props.format, src: props.src, file: props.file };
  const [load, setLoad] = useState(() => ({
    source,
    revision,
    documentClient: createOfficeDocumentClient(source, revision),
  }));
  if (
    load.source.name !== props.name ||
    load.source.format !== props.format ||
    load.source.src !== props.src ||
    load.source.file !== props.file ||
    load.revision !== revision
  ) {
    setLoad({ source, revision, documentClient: createOfficeDocumentClient(source, revision) });
  }
  const state = useSyncExternalStore(
    load.documentClient.subscribe,
    load.documentClient.getSnapshot,
    load.documentClient.getSnapshot,
  );
  if (state.status === "loading") return <FileSurfaceLoading />;
  if (state.status === "error")
    return (
      <FileSurfaceFailure
        message={state.message}
        onRetry={() => {
          props.onRetry?.();
          setRevision((current) => current + 1);
        }}
      />
    );
  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden">
      {state.document.warnings.map((warning) => (
        <FileSurfaceNotice key={warning}>{warning}</FileSurfaceNotice>
      ))}
      <DocumentContents key={`${props.name}:${props.src ?? "local"}`} document={state.document} />
    </div>
  );
}
