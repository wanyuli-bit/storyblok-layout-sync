# Storyblok Layout Sync

Copies the header and/or footer structure from a source market to other markets in Storyblok (space 336169). Every link is re-pointed to the target market's own page, and labels are set in the local language.

- **Header** follows **DK**'s layout.
- **Footer** follows **DE**'s layout. Each market keeps its own newsletter form, copyright line and sign-up texts.
- **Product links in the header** use the `title` of each market's product category page.

## Run it

Actions → **Storyblok Layout Sync** → Run workflow:

| Input | Notes |
|---|---|
| `storyblok_token` | Personal Access Token. Supplied fresh each run and masked in logs. Never stored as a secret. |
| `parts` | `header,footer`, `header` or `footer` |
| `markets` | e.g. `dk,no,se,fi`. The source market for a part is skipped automatically. |
| `mode` | `dry-run` (review only), `draft` (save, not published) or `publish` |

Start with `dry-run`, check the review file, then run `draft` and check in Storyblok before `publish`.

## Output

- **layout_sync_review.xlsx**: summary per market, every link with its target page and UUID, and any problems.
- **backups/**: each market's layout as it was before the run.

Both are attached to the run as an artifact. If `RESEND_API_KEY` and `MAIL_TO` secrets are set, the review file is also emailed.

If a link's target page doesn't exist in a market, that market isn't written and the run fails, so nothing broken goes in.

## Change labels or rules

All in `config/`:

- `header_labels.json` / `footer_labels.json`: labels per market, in the same order as the source layout. `@category_title` = take the title from the product category page.
- `link_rules.json`:
  - `remap` swaps part of a slug for a market (e.g. SE's DGNB link goes to Miljöbyggnad, localized Good Acoustics slugs).
  - `drop` removes items with no equivalent (e.g. Impressum everywhere, Mono Acoustic installers in FI).
- `settings.json`: space ID, source markets, and which footer fields to keep from the target market.

If DK's header or DE's footer changes structure (items added or removed), the run stops and tells you which column or category needs its labels updated.

To add a market, add its labels to both label files (and any rules), then include it in `markets`.

## Run locally

```bash
pip install -r requirements.txt
STORYBLOK_TOKEN=... python sync_layout.py --parts header,footer --markets no,se,fi --mode dry-run
```
