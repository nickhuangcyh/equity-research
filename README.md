# equity-research

Personal equity research — stock analysis, DCF models, earnings notes, and investment theses.

## Structure

```
reports/
  earnings/        # Post-earnings update notes
  initiations/     # Initiating coverage reports
  sector/          # Sector and industry overviews
  dcf/             # DCF valuation models and workbooks
  morning-notes/   # Daily morning meeting notes

theses/            # Investment theses and tracking
screens/           # Stock screening results and idea lists
```

## Setup

This workspace uses the [financial-services](https://github.com/nickhuangcyh/financial-services) agy plugin for research and modeling.

```bash
# Install plugin
git clone https://github.com/nickhuangcyh/financial-services /tmp/fs-plugin && \
agy plugin install /tmp/fs-plugin/agy-plugins/financial-services && \
rm -rf /tmp/fs-plugin
```

## Slash Commands

| Command | Description |
|---|---|
| `/dcf` | Build a DCF valuation model |
| `/comps` | Comparable company analysis |
| `/earnings` | Post-earnings update report |
| `/earnings-preview` | Pre-earnings scenario analysis |
| `/screen` | Stock screening and idea generation |
| `/sector` | Sector / industry landscape report |
| `/initiate` | Initiating coverage report |
| `/thesis` | Investment thesis tracker |
| `/catalysts` | Catalyst calendar |
| `/model-update` | Update model with new data |
| `/morning-note` | Morning meeting note |
