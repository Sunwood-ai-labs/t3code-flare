import type { EnvironmentId, ScopedThreadRef } from "@t3tools/contracts";
import type { OfficePreviewFormat } from "@t3tools/shared/filePreview";
import { lazy, Suspense, useEffect, useState } from "react";

import { useAssetUrlRefresh, useAssetUrlState } from "~/assets/assetUrls";
import { workspaceAssetResource } from "./filePreviewMode";
import { usePreviewAuthorizationVersion } from "./usePreviewAuthorizationVersion";
import { FileSurfaceFailure, FileSurfaceLoading } from "./fileSurfaceChrome";

const OfficeDocumentPreview = lazy(() => import("./OfficeDocumentPreview"));

/** Binary bytes come from the same authenticated, exact-file capability as other host previews. */
export function WorkspaceOfficePreview(props: {
  readonly environmentId: EnvironmentId;
  readonly threadRef: ScopedThreadRef;
  readonly draft: boolean;
  readonly absolutePath: string;
  readonly workspaceRoot: string;
  readonly format: OfficePreviewFormat;
  readonly workspaceMutationId: string | null;
  readonly reloadRevision?: number;
}) {
  const resource = workspaceAssetResource({
    kind: "media-file",
    threadRef: props.threadRef,
    draft: props.draft,
    workspaceRoot: props.workspaceRoot,
    absolutePath: props.absolutePath,
  });
  const refresh = useAssetUrlRefresh(props.environmentId, resource);
  const mutationId = props.reloadRevision
    ? `manual:${props.reloadRevision}:${props.workspaceMutationId ?? ""}`
    : props.workspaceMutationId;
  const [retryRevision, setRetryRevision] = useState(0);
  const requestKey = JSON.stringify([
    props.environmentId,
    props.threadRef.environmentId,
    props.threadRef.threadId,
    props.draft,
    props.workspaceRoot,
    props.absolutePath,
    mutationId,
    retryRevision,
  ]);
  const request = usePreviewAuthorizationVersion(requestKey, refresh);
  const requiresAuthorization = mutationId !== null || retryRevision > 0;
  const asset = useAssetUrlState(props.environmentId, requiresAuthorization ? null : resource);
  const [authorization, setAuthorization] = useState<{
    request: number;
    refresh: typeof refresh;
    url: string | null;
  } | null>(null);
  // oxlint-disable-next-line t3code/no-unreasoned-state-sync -- Gates the Office reader on a fresh server-issued exact-file capability for this request revision; cleanup ignores stale refresh completions.
  useEffect(() => {
    if (!requiresAuthorization) return;
    let cancelled = false;
    // Mint the revision's capability before mounting its binary reader. An old
    // cached capability can otherwise start a second download while this resolves.
    void refresh()
      .then((url) => {
        if (!cancelled) setAuthorization({ request, refresh, url });
      })
      .catch(() => {
        if (!cancelled) setAuthorization({ request, refresh, url: null });
      });
    return () => {
      cancelled = true;
    };
  }, [refresh, request, requiresAuthorization]);
  if (
    requiresAuthorization &&
    (authorization?.request !== request || authorization?.refresh !== refresh)
  )
    return <FileSurfaceLoading />;
  const url = requiresAuthorization
    ? authorization?.url
    : asset._tag === "Success"
      ? asset.url
      : undefined;
  if (url === null || (!requiresAuthorization && asset._tag === "Failure")) {
    return (
      <FileSurfaceFailure
        message="Unable to load this Office document. Open the original in another app."
        onRetry={() => setRetryRevision((revision) => revision + 1)}
      />
    );
  }
  if (url === undefined) return <FileSurfaceLoading />;
  const revision = `${props.workspaceMutationId ?? ""}:${props.reloadRevision ?? 0}`;
  const suffix =
    props.workspaceMutationId === null && !props.reloadRevision
      ? ""
      : `${url.includes("?") ? "&" : "?"}workspace-revision=${encodeURIComponent(revision)}`;
  return (
    <Suspense fallback={<FileSurfaceLoading />}>
      <OfficeDocumentPreview
        key={url + suffix}
        name={props.absolutePath}
        format={props.format}
        src={url + suffix}
        onRetry={() => setRetryRevision((revision) => revision + 1)}
      />
    </Suspense>
  );
}
