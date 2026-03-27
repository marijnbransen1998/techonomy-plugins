# Techonomy Plugins

Interne Claude Code plugin marketplace van Techonomy.

## Installatie (eenmalig)

```bash
/plugin marketplace add github:marijnbransen1998/techonomy-plugins
/plugin install techonomy-tools@techonomy-plugins
```

## Skills

### campagneplan-presentatie

Genereert een complete, Techonomy-branded campagneplan presentatie (.pptx) op basis van een client briefing.

Triggert automatisch wanneer je vraagt om een campagneplan, paid advertising strategy deck of mediaplan presentatie.

**Voorbeelden:**
```
"Maak een campagneplan presentatie voor KNVB, budget €25.000, kanalen Meta en Google Ads, periode april–mei 2026"
"Genereer een campagneplan deck voor client X"
"Maak een presentatie voor een paid advertising campagne"
```

De skill genereert een kant-en-klare .pptx met de Techonomy huisstijl, opgebouwd volgens de 4-chapter structuur:
**Inleiding → Campagne aanpak → Budget verdeling → Mediaplan**

## Updates

Wanneer er een nieuwe versie van een skill is, voer dan uit:
```bash
/plugin update techonomy-tools@techonomy-plugins
```

## Vereisten

- Claude Code CLI of Claude Desktop
- Python 3 met `python-pptx` (`pip3 install python-pptx`)
- Techonomy PPT template op: `/Users/<jouw-naam>/Downloads/Techonomy PPT met voorbeelden.pptx`
  _(pas het templatepad in `create_campagneplan.py` aan naar jouw eigen locatie)_
