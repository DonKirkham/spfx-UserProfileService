# Upgrade Plan: SPFx 1.10.0 → 1.23.0 (rebuild & port)

> **Status: COMPLETE.** The project was rebuilt on SPFx 1.23.0 (Heft, React 17, Fluent UI 8,
> PnPjs v4, TS 5.8) and builds end-to-end: `pnpm run build` produces
> `sharepoint/solution/user-profile-service.sppkg`. Done on branch `upgrade/spfx-1.23.0`.
>
> Notable deviations from the original plan:
> - Yeoman generator used for scaffolding (not the spfx-cli preview).
> - No project `.npmrc`; Node pinned via `.node-version` (fnm alias `spfx-1.23.0`).
> - pnpm 11 build-script approval lives in `pnpm-workspace.yaml`
>   (`allowBuilds: { unrs-resolver: true }`), not `package.json`.

Branch: `upgrade/spfx-1.23.0`
Approach: **scaffold a fresh SPFx 1.23.0 solution and port the existing code into it.**
Package manager: **pnpm**. Node: **v22** (via `fnm`).

## Why rebuild instead of in-place upgrade

The jump is 13 minor versions and crosses the **gulp → Heft toolchain replacement** (SPFx
1.22), Node 10→22, React 16→17, TS 3.3→5.8, Fabric 6 → Fluent UI 8, and PnPjs v2→v4.
Chasing that much config drift in place is more error-prone than generating a clean,
known-good 1.23.0 project and moving our small amount of source into it. This repo only
contains **one service and one React component** (plus stubbed/commented code), so porting
is cheap.

What we keep from the old repo:
- `src/services/UserProfileService.ts` (the real logic — ported, with PnPjs v4 changes)
- `src/webparts/ups/components/Ups.tsx` (the UI — ported, with Fluent UI 8 imports)
- web part identity from `UpsWebPart.manifest.json` (`preconfiguredEntries`, title, icon)
- `README.md`, `.git`, this plan

Everything else (gulp config, tsconfig, tslint, old `package.json`, `lib/`, `dist/`,
`config/*`) is **replaced** by the scaffolded output.

---

## Phase 0 — Environment

- [x] Branch `upgrade/spfx-1.23.0` created.
- [x] Node 22 active: `fnm use spfx-cli` (alias → v22.22.3; SPFx 1.23 requires Node 22 LTS).
      Note: the fnm alias is named `spfx-cli`, not `spfx-1.23.0`.
- [ ] Activate **pnpm** via corepack (corepack 0.34.6 is present):
      `corepack enable pnpm && corepack prepare pnpm@latest --activate`
- [ ] Install scaffolder. Two options:
  - **SPFx CLI (preview, recommended given the Node alias):** `pnpm add -g @microsoft/spfx-cli`
  - **Yeoman generator (stable):** `pnpm add -g yo @microsoft/generator-sharepoint@1.23.0`

### pnpm + SPFx caveat (important)

SPFx's build expects a hoisted/flat `node_modules`. pnpm's default isolated layout breaks
it (phantom-dependency resolution failures). Before installing deps, add a project
**`.npmrc`** with:

```ini
node-linker=hoisted
# (equivalently: shamefully-hoist=true)
```

Commit this `.npmrc` so every install is reproducible.

## Phase 1 — Scaffold a clean 1.23.0 React web part

Scaffold into a **temp directory** (outside the repo) so the generator runs against an empty
folder, then we merge the output in.

- [ ] Create temp dir, e.g. `~/tmp/ups-1230`.
- [ ] Scaffold matching the original identity:
  - **SPFx CLI:**
    `spfx create --template webpart-react --library-name user-profile-service --component-name "Ups"`
  - **Yeoman:** `yo @microsoft/sharepoint` → solution name `user-profile-service`,
    no tenant deploy / current site, framework **React**, web part name **Ups**.
- [ ] Choose **pnpm** as the package manager if the scaffolder prompts; otherwise delete the
      generated `package-lock.json` and run `pnpm install` after adding `.npmrc`.
- [ ] Confirm a clean baseline builds in the temp dir: `pnpm run build` (Heft) and
      `pnpm start` loads the local workbench.

## Phase 2 — Bring the scaffold into the repo

- [ ] From the temp solution, copy everything **except** `node_modules`, `.git`, and
      generated `src/webparts/ups/components/Ups.tsx` placeholder into the repo, overwriting:
      `package.json`, `tsconfig.json`, `.eslintrc.js`, `.npmignore`, `gulpfile.js` (will not
      exist — Heft), `config/*`, `.yo-rc.json`, `src/index.ts`, the web part scaffold.
- [ ] Delete obsolete repo files: `gulpfile.js`, `tslint.json`, `lib/`, `dist/`, `temp/`,
      old `config/*`, old `package-lock.json`.
- [ ] Keep `.git`, `README.md`, `UPGRADE-PLAN.md`, and the ported source (Phase 3).
- [ ] Add the `.npmrc` from Phase 0.
- [ ] `pnpm install` and confirm the scaffold-only project still builds inside the repo.

## Phase 3 — Port the source code

### `src/services/UserProfileService.ts` — port + PnPjs v2 → v4
- [ ] Copy the file in. Add PnPjs v4: `pnpm add @pnp/sp @pnp/core @pnp/logging` (v4.x).
- [ ] Replace the global `sp` usage with the `spfi()` factory and context:
  ```ts
  import { spfi, SPFx } from "@pnp/sp";
  import "@pnp/sp/profiles";

  // constructed with the web part context
  const sp = spfi().using(SPFx(this.context));
  const profile = await sp.profiles.myProperties();   // v4: no .get()
  ```
- [ ] Thread the web part `context` into the service (constructor/init param) — it can no
      longer rely on the old global `sp`.
- [ ] Keep `IUserProperty`, the AAD-property mapping, the `UserProfileProperties` merge, and
      `UserProfileServiceMock` (mock still useful for the local workbench).
- [ ] Delete the dead commented-out `profile.forEach` block.
- [ ] Fix `UserProfileServiceMock._profile: IUserProperty[] = null` for `strictNullChecks`
      (the Heft rig tsconfig is strict): use `IUserProperty[] | undefined` or `[]`.

### `src/webparts/ups/components/Ups.tsx` — port + Fluent UI 8
- [ ] Copy the file in. Change Fabric imports to Fluent UI 8:
  ```ts
  import { DetailsList, DetailsListLayoutMode, SelectionMode, List } from "@fluentui/react";
  ```
  (`office-ui-fabric-react` is gone; SPFx 1.23 ships `@fluentui/react` 8. The `DetailsList`
  props used here are API-compatible, so the JSX is unchanged.)
- [ ] Extend `IUpsProps` to receive `context` (and/or the service instance) so the real
      `UserProfileService` can be constructed with context.
- [ ] React 16→17 needs no source changes here.
- [ ] Optionally remove the dead commented `_renderList` block and unused `webpartState`
      scaffolding while touching the file (or leave as-is — out of scope).

### `src/webparts/ups/UpsWebPart.ts`
- [ ] Use the scaffolded 1.23 web part as the base; re-add rendering of the `Ups` component.
- [ ] Pass `this.context` into the `Ups` element props.

### `UpsWebPart.manifest.json`
- [ ] Use the scaffolded manifest (current `manifestVersion`/`$schema`); re-apply the
      original `preconfiguredEntries`: title `_UPS Demo`, description, icon `PeopleAlert`,
      group `Other`.

## Phase 4 — Build, lint, verify

- [ ] `pnpm run build` (Heft) — resolve TS 5.8 strictness + ESLint findings iteratively.
- [ ] `pnpm start` — local workbench renders the web part; mock service shows Property1–3
      in the `DetailsList`.
- [ ] Real SPO page workbench — confirm `sp.profiles.myProperties()` returns AAD + UPA
      properties for the signed-in user.
- [ ] `pnpm run build` / `heft package-solution --production` → produces `.sppkg`; upload to
      the App Catalog to validate.

## Phase 5 — Cleanup & commit

- [ ] Remove this leftover temp solution dir.
- [ ] Update `README.md` build steps (gulp → Heft / pnpm; the README's `gulp ...` commands
      are now wrong).
- [ ] Commit on `upgrade/spfx-1.23.0`; open PR.

## Risks & notes

- **pnpm hoisting** is the most likely early blocker — set `.npmrc node-linker=hoisted`
  before the first install.
- **PnPjs v2→v4** is the only nontrivial code change: it requires plumbing `context` through
  `UpsWebPart → Ups → UserProfileService` and dropping `.get()`.
- **Hosted workbench is deprecated** (retires Dec 1 2026) — use the local workbench / SPFx
  Debug Toolbar for testing.
- The stubbed buttons / `WebpartState` enum in `Ups.tsx` are unimplemented; not part of this
  upgrade.
- SPFx CLI is still **preview** (GA targeted June 2026). If it misbehaves, fall back to the
  Yeoman generator 1.23.0 — the rest of the plan is identical.

## Reference links

- [SPFx 1.23.0 release notes](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/release-1.23.0)
- [SPFx CLI (preview)](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/sharepoint-framework-cli)
- [Heft toolchain](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/sharepoint-framework-toolchain-rushstack-heft)
- [SPFx compatibility matrix](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/compatibility)
- [PnPjs getting started (spfi / SPFx)](https://pnp.github.io/pnpjs/getting-started/)
