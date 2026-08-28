# Rasuwa Flood

Source code for [rasuwaflood.org](https://rasuwaflood.org), a public, source-based record of the 26 August 2026 Rasuwa–Bhote Koshi–Trishuli flood.

## Local preview

Run `python3 -m http.server 8000`, then open `http://localhost:8000`.

## Publishing

GitHub Pages publishes directly from the root of `main`. The `CNAME` file assigns the custom domain. DNS must point the apex domain to GitHub Pages before HTTPS can be enabled.

## Automated source discovery

GitHub Actions runs `scripts/update_sources.py` every three hours and on demand. The collector reads news-search feeds, applies event-specific keyword checks, deduplicates URLs, and updates `data/latest.json`. It indexes titles, links and short feed descriptions only; it does not republish articles or treat automated discovery as verification.

## Editorial note

Event information is provisional. Update `data/latest.json` and the source record in `index.html` as authoritative assessments develop. Confirm that publication rights cover every locally hosted image.
