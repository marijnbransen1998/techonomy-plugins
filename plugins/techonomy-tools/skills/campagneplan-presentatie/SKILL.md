---
name: campagneplan-presentatie
description: Generates a complete Techonomy-branded campaign plan PowerPoint presentation (.pptx) based on a client briefing. Use this skill whenever the user asks to make a campaign plan presentation, a campagneplan, a paid advertising strategy deck, a media plan presentation, or wants to generate a PowerPoint for a client campaign. Also trigger when the user provides campaign details like client name, budget, channels, or target audience and mentions a presentation or deck. The skill produces a ready-to-present .pptx file following Techonomy's exact branding, structure, and 4-chapter narrative arc (Inleiding → Campagne aanpak → Budget verdeling → Mediaplan).
---

# Campagneplan Presentatie

Generates a complete, Techonomy-branded campaign plan presentation in PowerPoint format using the official Techonomy template.

## When to use

Use this skill whenever the user:
- Asks to make / generate / create a campagneplan presentatie
- Provides a campaign briefing and wants a slide deck
- Asks for a paid advertising strategy presentation
- Mentions a new client campaign that needs a deck

## Brand & Template

- **Template**: `/Users/marijnbransen/Downloads/Techonomy PPT met voorbeelden.pptx`
- **Heading font**: Orbitron Bold
- **Body font**: Montserrat
- **Colors**: Dark navy `#000031`, Dark blue `#0000DC`, Orange-red `#FF4600`, Pink `#D4145A`, Light gray `#E6EEEF`, Light blue `#94CEE5`
- **Script**: `scripts/create_campagneplan.py` (use this — do not rewrite from scratch)

## Presentation Structure

Every campagneplan follows this exact 4-chapter arc:

```
[Cover]         Campagnenaam | Client × Techonomy | Datum
[Inhoudsopgave] 01 Inleiding / 02 Campagne aanpak / 03 Budget / 04 Mediaplan
─── H1: Inleiding & debrief ──────────────────────────────────────────────────
  Briefing slide: Aanleiding, Doelstelling, Doelgroep, Campagneperiode, Budget
─── H2: Campagne aanpak ──────────────────────────────────────────────────────
  Tijdlijn       → visuele campagnetijdlijn (oplevering → start → einde → eval)
  Kanalen        → tabel: Fase | Kanaal | Doelstelling | Looptijd
  Doelgroepen    → targeting per platform (Meta / Google / TikTok)
  Copies         → advertentietekst-suggesties per fase
  Assets         → gewenste creatieve formaten per platform
─── H3: Budget verdeling ─────────────────────────────────────────────────────
  Budget split   → Awareness / Verkeer / Conversie + 2,5% admin fee
─── H4: Mediaplan ────────────────────────────────────────────────────────────
  Awareness      → Kanaal | Impressies | CPM | Budget
  Verkeer        → Kanaal | Impressies | Klikken | CTR | CPC | Budget
  Conversie      → Kanaal | Impressies | CPA | Budget | Conversies
[Logo]          Afsluitende dia
```

## Kanalen slide design (confirmed spec — do not deviate)

The "Selectie campagnekanalen" slide uses a table with the same base styling as the mediaplan, plus platform logos in the Kanaal column:

- **Header row**: `#010031` background, Orbitron Bold White 10pt — identical to mediaplan header
- **Fase column**: funnel-phase color per row (Awareness=`#0000DC`, Verkeer=`#FF4600`, Conversie=`#D4145A`), Orbitron Bold White 9pt
- **Kanaal column**: `#1600DC` background, Orbitron Bold White 8pt + platform logos as floating image overlays:
  - Meta rows → Facebook + Instagram logos side by side
  - Google Display / YouTube → Google Ads + YouTube logos
  - Google Search / Performance Max → Google Ads logo only
  - TikTok rows → TikTok logo
- **Doelstelling / Looptijd columns**: white background, Montserrat 9pt dark navy
- Row height is fixed at 0.52" so logo Y positions can be calculated precisely
- Logo assets: `assets/icons/logo_facebook.png`, `logo_instagram.png`, `logo_google_ads.png`, `logo_tiktok.png`, `logo_youtube.png`
- Logo size: 0.30" × 0.30", vertically centered in the row, placed left of the channel name text

## Mediaplan table design (confirmed spec — do not deviate)

The mediaplan tables must use these exact styles:
- **Header row background**: `#010031` (near-black navy) — `RGBColor(1, 0, 49)`
- **Header row font**: Orbitron Bold White, 10pt
- **Kanaal column cells (data rows)**: background `#1600DC` (bright blue) — `RGBColor(22, 0, 220)`
- **Kanaal column font**: Orbitron Bold White
- **All other data cells**: white background, Montserrat 8pt, dark navy text
- **Section label**: "Opmerking:" (not "Toelichting:")
- Column headers for Awareness: Kanaal | Impressies | CPM | Budget
- Column headers for Verkeer: Kanaal | Impressies | Klikken | CTR | CPC | Budget
- Column headers for Conversie: Kanaal | Impressies | CPA | Budget | Conversies

## Timeline design (confirmed spec — Techonomy style)

The campagne tijdlijn slide uses:
- Horizontal dark-navy bar running the full content width at vertical center
- 5 steps: Oplevering campagneplan → Voorbereiding & assets → Start paid campagne → Einde campagne → Eindevaluatie
- Each step has a **dark-blue filled circle** sitting on the bar with a white step number (Orbitron Bold, 9pt)
- A short vertical **connector stub** (dark navy) links the circle to the label box
- Steps alternate **above** (odd: 1,3,5) and **below** (even: 2,4) the timeline bar
- Label boxes: rounded rectangles, light gray fill (`#E6EEEF`), dark navy bold text 7pt, date in dark blue 7pt beneath
- Start and end dates from config are shown on steps 3 and 4 respectively

## Funnel model (always use this)

**Awareness → Verkeer → Conversie**
Default budget split: 40% / 30% / 30%

## Standard platforms

| Platform | Always? | Phase |
|---|---|---|
| Meta (Facebook/Instagram) | Yes | Awareness + Verkeer + Conversie |
| Google Ads (Search/PMax/Display) | Yes | Awareness + Verkeer + Conversie |
| TikTok | Optional | Awareness |
| Programmatic Display | Optional | Awareness |

---

## Step-by-step workflow

### 1. Gather briefing information

Ask the user for (or extract from their message):

| Field | Required | Default |
|---|---|---|
| `client` | Yes | — |
| `campaign_name` | Yes | — |
| `start_date` | Yes | — |
| `end_date` | Yes | — |
| `total_budget` | Yes | — |
| `target_audience` | Yes | — |
| `channels` | No | `["Meta", "Google Ads"]` |
| `campaign_notes` | No | `""` |
| `awareness_pct` | No | `40` |
| `verkeer_pct` | No | `30` |
| `conversie_pct` | No | `30` |
| `meta_geo`, `meta_gender`, `meta_age`, `meta_targeting` | No | see defaults |
| `google_geo`, `google_gender`, `google_age`, `google_targeting` | No | see defaults |
| `tiktok_geo`, `tiktok_gender`, `tiktok_age`, `tiktok_targeting` | No | see defaults |

If the user provides all of this in their initial message, skip directly to step 2. If key fields are missing, ask once in a single message — don't ask field by field.

### 1b. Think about the campaign — what else should go in the deck?

Before building the config, **reason briefly about the briefing**. The fixed structure above is a minimum. Based on what you know about the campaign, the client, and the sector, consider adding extra slides to the deck. These should be added AFTER generating the base PPTX, as additional custom slides written directly in python-pptx using the `Alleen titel - 1` layout + textboxes — or by appending them to the script run.

**Ask yourself:**
- Is this a seasonal campaign? → Add a seasonal timing / calendar slide explaining why now.
- Is the budget split unusual (e.g. 90% conversie)? → Add a rationale slide.
- Is the client new to paid advertising? → Add an introductory slide explaining the funnel model.
- Are there A/B test variants or creative hypotheses? → Add a creative strategy slide.
- Does the campaign have a specific KPI (e.g. CPA target, ROAS goal)? → Add a KPI & succes metrics slide.
- Is there a retargeting strategy beyond standard? → Add a retargeting flow slide.
- Does the sector have specific compliance rules (healthcare, gambling, alcohol)? → Add a compliance notes slide.
- Are there specific creative assets the client needs to produce? → Expand the assets slide with deadlines and responsibilities.
- Is there a phased rollout (e.g. teaser → launch → sustain)? → Add a phased campaign strategy slide.

You don't need to add all of these — pick 1-3 that are genuinely relevant given the briefing. Briefly tell the user which extra slides you're adding and why, before running the script.

### 2. Create config JSON

Write a `campagneplan_config.json` in the current working directory with all collected values. Example:

```json
{
  "client": "KNVB",
  "campaign_name": "Ticketverkoop NL – IRL",
  "start_date": "01-04-2026",
  "end_date": "15-05-2026",
  "total_budget": 25000,
  "target_audience": "Voetbalfans 18-55, Nederland",
  "channels": ["Meta", "Google Ads", "TikTok"],
  "campaign_notes": "Kaartverkoop voor interland NL-IRL, focus op conversie",
  "awareness_pct": 35,
  "verkeer_pct": 30,
  "conversie_pct": 35
}
```

### 3. Install dependencies and run script

```bash
pip install python-pptx -q
python scripts/create_campagneplan.py campagneplan_config.json campagneplan_[client].pptx
```

The script path is relative to the skill directory:
`/Users/marijnbransen/.claude/plugins/marketplaces/claude-plugins-official/plugins/techonomy-tools/skills/campagneplan-presentatie/scripts/create_campagneplan.py`

### 4. Report output

Tell the user the exact file path of the generated .pptx so they can open it. If the script produces any warnings about missing layouts, mention them so the user can verify the result in PowerPoint.

---

## Google Slides compatibility (critical)

All slides must render correctly in Google Slides, which does **not** support PowerPoint's auto-shrink. Key rules the script enforces — never break these:

- `tf.auto_size = MSO_AUTO_SIZE.NONE` on **every** text frame (including table cells and placeholders)
- Explicit `Pt(size)` on every run — never rely on the template's default font size
- Use `_force_tf(tf, size, ...)` after setting placeholder text to override every run's font size
- Chapter divider placeholders: `_force_tf` with 38pt (title) and 52pt (number) — the template placeholder has huge auto-shrink defaults that GS ignores

## Asset inventory

Icons and logos stored in `assets/icons/`:

| File | Used on | Description |
|---|---|---|
| `icon_location.png` | Doelgroepen | Location pin targeting icon |
| `icon_gender.png` | Doelgroepen | Gender targeting icon |
| `icon_age.png` | Doelgroepen | Age/calendar targeting icon |
| `icon_interests.png` | Doelgroepen | Interests/heart targeting icon |
| `logo_facebook.png` | Kanalen | Facebook logo |
| `logo_instagram.png` | Kanalen | Instagram logo |
| `logo_google_ads.png` | Kanalen | Google Ads logo |
| `logo_tiktok.png` | Kanalen | TikTok logo |
| `logo_youtube.png` | Kanalen | YouTube logo |

All icons were extracted from the official Techonomy template (slides 59/60) and example campagneplannen. The script uses `add_picture_safe()` which silently skips missing files and logs a warning.

## Notes on the template

- All slide layouts are loaded from the Techonomy template master — colors and fonts are theme-based, not hardcoded per shape.
- Layout names used: `Titeldia met afbeeldingsveld` (cover), `Alleen titel - 1` (all content slides), `Hoofdstuk met afbeeldingsveld` (chapter dividers), `Logo` (closing slide)
- If a named layout is not found, the script falls back to the first available layout and prints a warning.
- After generating, the user may want to add client photos/images to the image placeholder fields manually in PowerPoint.
- The template path is fixed at `/Users/marijnbransen/Downloads/Techonomy PPT met voorbeelden.pptx` — if the user moves the file, update this path.
