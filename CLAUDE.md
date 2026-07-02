# Claudecode — Repo Map & Rules

This directory holds several **unrelated** personal projects side by side. Read this before
moving, deleting, or assuming anything about what's "dead" — a lot of it looks like clutter
but has an active deployment behind it.

## Deployment topology (this is the part that bites)

There are **three separate Railway deployments** sourced from **two separate git repos**, plus
this main repo itself:

| Deployment | Repo | Source path | Live URL |
|---|---|---|---|
| OutsideFramework (portfolio homepage) | this repo (`origin` = `global-economic-globe.git`) | `app/OutsideFramework/index.html` → `Dockerfile` → nginx | — |
| globe / invest / causal / brownian | **`globe-invest/` — its own repo** (`globe-invest.git`) | `globe-invest/app/{globe,invest,causal,brownian}/index.html` + `globe-invest/server.js` | `globe-invest.up.railway.app` |
| structural-holes | **`structural_holes/` — its own repo** | — | `structural-holes-production.up.railway.app` |
| article-db | this repo, `article_db/` (already tracked) | — | `articlebase.up.railway.app` |
| my-slide | **`my-slide/` — its own repo** (Netlify) | — | — |

**`app/GlobalEco`, `app/InvestFrame`, `app/CausalFrame` are dev-source copies only — the
Dockerfile above does NOT deploy them.** The versions that actually go live are the mirrored
copies inside `globe-invest/app/`. Editing one side and forgetting the other is a real,
already-happened bug: an unrelated "fix sea routes" commit once silently dropped a commodity
and mislabeled another in GlobalEco's oil-price panel, and it went unnoticed for ~2 weeks
because nothing diffed the two repos.

**Whenever you edit GlobalEco / InvestFrame / CausalFrame: run
`scripts/sync-globe-invest.ps1`** to copy the change into `globe-invest/app/`, then commit +
push **both** repos (`origin` here, and `globe-invest`'s own `origin`). Don't rely on memory to
keep them in sync — the script diffs before copying.

### Known divergence — do not blind-sync these
`globe-invest/app/options`, `/warning`, and `/research` also started as copies of root-level
files (`options_guide.html` [now `projects/options_guide.html`], the old warning-radar HTML
[now `projects/market-warning-radar/`], and the MU analysis report [now
`research/mu-analysis-2026q3/`]). As of 2026-07-01, `options` is still identical; `warning`
and `research` have diverged into genuinely different content (different structure/theme, not
just "a few commits behind"). Don't add these to the sync script or overwrite either side until
a human reconciles which version is current.

### Another one: projects/brownian-motion-simulator
`globe-invest/app/brownian/index.html` (added 2026-07-01, deployed as the "Brownian Motion" work
on the OutsideFramework Works page) is a copy of `projects/brownian-motion-simulator/index.html`.
Same rule as above — it's **not** in the sync script. Edit the `projects/` copy for dev/testing,
then manually re-copy to `globe-invest/app/brownian/index.html` if it changes, and commit + push
both repos.

## Directory map

- `app/` — dev-source for the four "outside framework" apps (see table above)
- `globe-invest/`, `my-slide/`, `structural_holes/` — **separate git repos**, nested here for
  convenience. Never `git add` their contents into this repo; they manage their own history.
- `article_db/` — tracked in this repo, deployed standalone to `articlebase.up.railway.app`
- `projects/` — standalone tools/apps not part of the outside-framework family (stock analyzer,
  tech value screener, food calorie lookup, market warning radar, options guide, brownian motion
  simulator — the last one is also deployed via `globe-invest/app/brownian`, see above)
- `research/` — one-off analysis writeups (e.g. earnings-call breakdowns), not living apps
- `archive/` — superseded/legacy material kept for reference, not maintained
- `RR4/`, `RR5/`, `EMA Bias Model/` — active personal trading/VBA projects, not part of the web
  app suite. Left at the repo root deliberately — do not reorganize without asking.
- `scripts/` — maintenance scripts (currently: `sync-globe-invest.ps1`)

## Commit hygiene

Keep commits scoped to one concern. Several regressions in this repo's history came from
otherwise-correct commits that also carried an unrelated, unreviewed change (e.g. a routing-fix
commit that happened to rewrite the oil-price panel and dropped a field). If a change touches an
unrelated file, split it into its own commit even when working fast.

## Workflow preference

This repo is iterated on fast and pushed directly to `main` on both `origin` and `globe-invest` —
that's a deliberate choice for solo-project velocity, not an oversight. Don't propose branch
protection or PR gating; instead lean on the sync script above and `globe-invest`'s CI
(`.github/workflows/ci.yml`, runs `node --check` on push) as lightweight safety nets that don't
block pushes.
