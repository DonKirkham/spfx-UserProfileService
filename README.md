## spfx-UserProfileService

Quick SPFx web part that demonstrates a `UserProfileService` for SharePoint Online. It reads
the current user's Azure AD properties and SharePoint User Profile properties (via PnPjs) and
renders them in a Fluent UI `DetailsList`. Outside SharePoint (e.g. the local workbench) it
falls back to a mock service.

Built on **SPFx 1.24.1** (Heft toolchain, React 18, Fluent UI 8, PnPjs v4, TypeScript 5.8).

### Prerequisites

* Node.js **22, 24 or 26** — this repo pins the version via `.node-version` (fnm alias `spfx-1.24.1`, Node 26).
  With [fnm](https://github.com/Schniz/fnm): `fnm use` (reads `.node-version`).
* npm as the package manager.

### Building the code

```bash
git clone <repo>
fnm use          # selects Node 26 from .node-version
npm install
npm run build    # heft test --production && heft package-solution --production
```

The `.sppkg` is produced at `sharepoint/solution/user-profile-explorer.sppkg` — upload it to
your tenant App Catalog.

### Build options (Heft)

| Command          | Purpose                                              |
| ---------------- | ---------------------------------------------------- |
| `npm start`      | `heft start --clean` — local workbench dev server    |
| `npm run build`  | production test + bundle + package the `.sppkg`      |
| `npm run clean`  | `heft clean` — remove build output                   |
| `npx heft build` | compile + lint + webpack bundle (no package)         |
