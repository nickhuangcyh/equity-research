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
# Run from this workspace root
git clone https://github.com/nickhuangcyh/financial-services /tmp/fs && \
bash /tmp/fs/install-agy-local.sh && \
rm -rf /tmp/fs
```

Then restart `agy` in this directory.

## Slash Commands

| Command | Description |
|---|---|
| `/financial-analysis:dcf-model` | Build a DCF valuation model |
| `/financial-analysis:comps-analysis` | Comparable company analysis |
| `/financial-analysis:lbo-model` | LBO model |
| `/financial-analysis:competitive-analysis` | Competitive landscape analysis |
| `/financial-analysis:audit-xls` | Excel model audit |
| `/equity-research:earnings-analysis` | Post-earnings update report |
| `/equity-research:earnings-preview` | Pre-earnings scenario analysis |
| `/equity-research:idea-generation` | Stock screening and idea generation |
| `/equity-research:sector-overview` | Sector / industry landscape report |
| `/equity-research:initiating-coverage` | Initiating coverage report |
| `/equity-research:thesis-tracker` | Investment thesis tracker |
| `/equity-research:catalyst-calendar` | Catalyst calendar |
| `/equity-research:model-update` | Update model with new data |
| `/equity-research:morning-note` | Morning meeting note |
