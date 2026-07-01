# Releasing

`@mitresthen/excelents` publishes to npm automatically when a **GitHub Release**
is published. Auth is npm **Trusted Publishing (OIDC)** — there is no `NPM_TOKEN`
secret to manage.

## Cutting a release

1. Be on `master` with a clean tree and CI green.
2. Bump the version — this also rewrites the `version` export in `src/index.ts`
   (via the `version` npm script → `scripts/sync-version.mjs`) and commits + tags:

   ```sh
   npm version patch   # or: minor | major
   ```

3. Push the commit and the tag:

   ```sh
   git push --follow-tags
   ```

4. Create a GitHub Release for the new `vX.Y.Z` tag
   (`gh release create vX.Y.Z --generate-notes`, or the web UI).

Publishing the release triggers `.github/workflows/publish.yml`, which reinstalls,
verifies the tag matches `package.json`, runs the tests + build, and
`npm publish`es with provenance. Watch the run under the repo's **Actions** tab.

## One-time npm setup (already done for this package)

- The npm org `mitresthen` exists and the package name is published once (a
  trusted publisher can't be configured on a package that doesn't exist yet).
- On npmjs.com → the package → **Settings → Trusted Publishing**, add a GitHub
  Actions publisher pointing at this repo and `.github/workflows/publish.yml`.

Once that's in place, every published GitHub Release ships a new version.
