# Maintaining T3 Code Flare

Keep `main` as the maintained fork branch and `upstream` pointed at `https://github.com/pingdotgg/t3code.git`. Preserve merge history so upstream updates retain the fork's branding, artwork, and Office preview changes.

1. Fetch `upstream`, record the selected upstream commit, and create a fresh `sync/upstream-YYYYMMDD` branch from fork `main` in an isolated checkout.
2. Merge the recorded commit. Resolve source conflicts without replacing the fork's features. Inspect changed workflows and keep upstream deployment, signing, and publishing jobs restricted to `pingdotgg/t3code`.
3. Install from the frozen lockfile. Run focused tests for affected code, including artwork/settings and Office/parser/asset behavior, plus affected type and lint checks. Confirm both features in a real client with isolated T3 home data and save screenshots outside the repository.
4. Push the `sync/**` branch to this fork. Wait for the canonical GitHub CI `Check` job to pass on that exact head. Fix failures on the same branch; do not reuse results from another commit.
5. Advance fork `main` through a normal fast-forward only after review, CI, and client checks pass. If `main` has moved, merge it into the candidate and repeat affected verification. Do not force push.

The fork's CI uses GitHub-hosted runners. Inherited non-CI workflow jobs are restricted to the upstream repository. Audit newly added workflows during every upstream merge. This repository does not configure deployment, package publication, signing, or a separate installer/update channel for Flare.
