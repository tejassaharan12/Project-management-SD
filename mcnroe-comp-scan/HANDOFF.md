# Mcnroe competitor scan: handoff

Pick-up brief for the next session. Read this first, then continue the scan.

## Docs (Claude Docs, edit these, do not recreate)

- Tools and requirements: https://claude.ai/code/artifact/82165b92-3252-4d8f-8af4-1b6f672f27ab
- Competitor scan, Jun to Oct 2026 (press layer done): https://claude.ai/code/artifact/b7307ea6-30ea-466f-9f1c-9db21c95589a
- Earlier proposal (AM team owns it now): https://claude.ai/code/artifact/04b46bb4-fe99-4ef4-847d-fc30032ee901

## Brands (12)

Denver, BlaBliBlu, Bella Vita (Bellavita), Indē Wild, Secret Alchemist, Beardo, Park Avenue (Godrej),
Engage (ITC), Ajmal, EM5 (House of EM5), Fraganote, Plum (incl. BodyLovin').

## Scope still to do

Window: 15 Jul to 15 Oct 2026 (include June launches still running). Out of scope: influencers, marketplaces.

| Layer | Source | Tool |
| --- | --- | --- |
| Meta ads | Meta Ad Library, country IN, all media | Apify `apify/facebook-ads-scraper` |
| Google Search, YouTube, Display ads | Google Ads Transparency Center, region IN | Apify `solidcode/ads-transparency-scraper` |
| Instagram organic | Brand handles, posts newer than window start, comments on top posts | Apify `apify/instagram-scraper` |
| YouTube organic | Brand channels, uploads in window, views, likes, comments | YouTube Data API v3 |
| Google Trends | Brand names, India, past 12 months and window, 5 per comparison, anchor brand in every set | Apify `apify/google-trends-scraper` |

Env vars expected: `APIFY_TOKEN`, `YOUTUBE_API_KEY`. Network: full access granted.

Steps:

1. Find and verify each brand's Instagram handle, YouTube channel, Facebook page and Google advertiser name.
2. Run each layer, save raw exports to `mcnroe-comp-scan/data/`.
3. Code ads and posts on the framework lenses (territory, claim, format, talent, offer, response).
4. Add one section per layer to the scan doc, update the headline and brand snapshots, spot-check press dates.
5. Record actual Apify spend per layer and update the cost section in the tools doc.

## Known limits

- Meta Ad Library in India shows only ads still running; ads stopped before the scan cannot be recovered.
- No source gives competitor ad spend; report creative volume and run length instead.
