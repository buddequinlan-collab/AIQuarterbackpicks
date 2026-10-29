# AI Quarterback Picks

Static site for posting and selling model-priced picks (NFL, NBA, college football, NHL, MLB).

- `site/index.html` is the whole site. `site/data.js` holds the picks and track record it shows.
- One pick per card is free. The rest are written to `data.js` as locked rows (game and market only), so premium picks never appear in the public source. Members get the full card through Whop.
- Whop checkout link and price are set in `SETTINGS` near the bottom of `site/index.html`.

## Updating picks

`scripts/export_site_data.py` rebuilds `site/data.js` from the slate workbooks (the "Named picks" sheet of each `projections/*.xlsx`) and the graded `results/bet_tracker.xlsx`:

```
PROJECT_FILES=/path/to/project-files python3 scripts/export_site_data.py
```

Then commit and push `site/data.js`; the host redeploys on push.

## Hosting

Cloudflare Pages (or Netlify) connected to this repo: no build command, output directory `site`.

For entertainment only, not betting advice. 21+. Problem gambling help: 1-800-GAMBLER.
