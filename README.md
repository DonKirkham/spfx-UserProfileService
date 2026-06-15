## spfx-UserProfileService

Quick SPFx web part that demonstrates a `UserProfileService` for SharePoint Online. It reads
the current user's Azure AD properties and SharePoint User Profile properties (via PnPjs) and
renders them in a Fluent UI `DetailsList`. Outside SharePoint (e.g. the local workbench) it
falls back to a mock service.

Built on **SPFx 1.23.0** (Heft toolchain, React 17, Fluent UI 8, PnPjs v4, TypeScript 5.8).

### Prerequisites

* Node.js **22 LTS** — this repo pins the version via `.node-version` (fnm alias `spfx-1.23.0`).
  With [fnm](https://github.com/Schniz/fnm): `fnm use` (reads `.node-version`).
* [pnpm](https://pnpm.io) as the package manager.

### Building the code

```bash
git clone <repo>
fnm use          # selects Node 22 from .node-version
pnpm install
pnpm run build   # heft test --production && heft package-solution --production
```

The `.sppkg` is produced at `sharepoint/solution/user-profile-service.sppkg` — upload it to
your tenant App Catalog.

### Build options (Heft)

| Command                | Purpose                                              |
| ---------------------- | ---------------------------------------------------- |
| `pnpm start`           | `heft start --clean` — local workbench dev server    |
| `pnpm run build`       | production test + bundle + package the `.sppkg`      |
| `pnpm run clean`       | `heft clean` — remove build output                   |
| `pnpm exec heft build` | compile + lint + webpack bundle (no package)         |
