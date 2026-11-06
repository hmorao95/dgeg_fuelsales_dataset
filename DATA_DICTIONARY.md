# Data dictionary — DGEG fuel sales

Applies to `outputs/oil_products_panel.csv`, `outputs/oil_products_wide.*`,
`outputs/oil_products_timeseries.*` and `outputs/oil_products_timeseries_sa.*`.

## Panel file (`oil_products_panel.csv`)

Long format: one row per product × month (13,152 rows, 50 products).

| Column | Description |
|---|---|
| `period` | Month as `YYYYmM`, e.g. `1995m1`. |
| `date_monthly` | First day of the month, ISO date `YYYY-MM-01`. |
| `year`, `month` | Calendar year and month. |
| `product_id` | Numeric product code 1–50, in alphabetical order of `product`. |
| `product` | Product name, as in the wide files' column headers. |
| `consumption` | Sales in tonnes; missing where the product was listed but not reported. |

Panel index: (`product_id`, `date_monthly`).

## Wide and time-series files

### Time columns

| Column | Description |
|---|---|
| `period` | Month as `YYYYmM`, e.g. `1995m1`, `2026m9`. |
| `date_monthly` | First day of the month, ISO date `YYYY-MM-01` (wide files only). |

Coverage: **1995m1 – 2026m9**, monthly. `oil_products_timeseries.*` has a
gap-free monthly index; blank cells mean the product had no reported sales that month.

## Seasonally adjusted file (`oil_products_timeseries_sa.*`)

Same layout as `oil_products_timeseries.*` (`period` + 50 product columns, tonnes),
with each product's sales seasonally adjusted:

- **Method:** STL (Cleveland et al., 1990) on log sales, 12-month period,
  13-term seasonal smoother, robust weights against one-off spikes. Adjusted
  sales = sales ÷ seasonal factor, with the factors rescaled so any 12
  consecutive months average 1 (as in X-11), so the adjusted series keeps the
  original's level. Values rounded to 0.1 t.
- **Not adjusted (blank column):** products with less than 36 months of data, a
  gap, or a month at zero, and products whose adjusted annual totals differ from
  the original's by more than 5% (median over full years). With data to
  2026m9, 31 of 50 products are adjusted.
- **Not included:** trading-day and Easter effects. These are left in the series.
- **Revisions:** the whole history is re-estimated at every update, so past
  adjusted values change slightly when DGEG publishes a new month.
- Aggregates (ⓐ) are adjusted directly, not as the sum of adjusted components,
  so they need not add up exactly.

## Units & conventions

- **Units:** metric tonnes (t) — DGEG files state "unidade: tonelada".
- **BTE** = *baixo teor de enxofre* (low-sulphur); **ATE** = *alto teor de enxofre* (high-sulphur).
- **`marítimo`** = marine/bunker variant of the product.
- Blank/empty = not reported (not zero). Genuine zeros are stored as `0`.
- Some columns are **aggregate totals or memo lines** (marked ⓐ below) — do **not** sum them together with their components, or you will double-count.

## Product columns (50)

### Internal market — LPG & gases
| Column | English | Notes |
|---|---|---|
| `Butano` | Butane | LPG |
| `Propano` | Propane | LPG |
| `Gás Auto` | Auto LPG (autogas) | road transport LPG |
| `Propileno` | Propylene | petrochemical feedstock |

### Internal market — gasolines
| Column | English | Notes |
|---|---|---|
| `Gasolina s/Chumbo 95` | Unleaded gasoline 95 RON | |
| `Gasolina s/Chumbo 98` | Unleaded gasoline 98 RON | |
| `Gasolina super aditivada` | Additised "super" gasoline | leaded-petrol substitute, older years |

### Internal market — diesel / gas oils
| Column | English | Notes |
|---|---|---|
| `Gasóleo rodoviário` | Road diesel | |
| `Gasóleo aquecimento` | Heating gas oil | |
| `Gasóleo colorido e marcado` | Coloured & marked gas oil | reduced-tax off-road/agriculture |

### Internal market — kerosenes
| Column | English | Notes |
|---|---|---|
| `Petróleo iluminante` | Illuminating kerosene (lamp oil) | |
| `Petróleo carburante` | Power/vaporising kerosene | |
| `Petróleos (iluminante + carburante)` ⓐ | Kerosenes total | aggregate of the two above |

### Internal market — fuel oils
| Column | English | Notes |
|---|---|---|
| `Fuelóleo nº3 BTE` | Fuel oil no.3, low-sulphur | |
| `Fuelóleo nº4 BTE` | Fuel oil no.4, low-sulphur | |
| `Fuelóleo nº4 ATE` | Fuel oil no.4, high-sulphur | |

### Memo — fuel oil by use
| Column | English | Notes |
|---|---|---|
| `Fuelóleo para a produção de energia elétrica` ⓐ | Fuel oil for electricity generation | memo line |
| `Fuelóleo para os restantes setores` ⓐ | Fuel oil for other sectors | memo line |

### Internal market — other products
| Column | English | Notes |
|---|---|---|
| `Coque de petróleo` | Petroleum coke (petcoke) | |
| `Asfaltos` | Asphalt / bitumen | |
| `Lubrificantes` | Lubricants | |
| `Parafinas` | Paraffins / waxes | |
| `Solventes` | Solvents | |
| `Nafta Química` | Chemical naphtha | |
| `Nafta Química e Aromáticos` ⓐ | Chemical naphtha & aromatics | aggregate |
| `Matéria Prima de Aromáticos` | Aromatics feedstock | |
| `Burner's Oil` | Burner's oil | heating distillate |
| `Resíduo AV` | Vacuum/atmospheric residue | heavy refinery residue |

### Marine bunker market (`marítimo`)
| Column | English | Notes |
|---|---|---|
| `Mercado de Bancas Marítimas` ⓐ | Marine bunker market | aggregate total |
| `Mercado de Bancas` ⓐ | Bunker market | aggregate total |
| `Gasóleo marítimo` | Marine gas oil | |
| `Gasóleo colorido e marcado marítimo` | Coloured & marked gas oil, marine | |
| `Fuelóleo nº3 BTE marítimo` | Fuel oil no.3 LS, marine | |
| `Fuelóleo nº4 BTE marítimo` | Fuel oil no.4 LS, marine | |
| `Fuelóleo nº4 ATE marítimo` | Fuel oil no.4 HS, marine | |
| `Lubrificantes marítimo` | Lubricants, marine | |

### Aviation market
| Column | English | Notes |
|---|---|---|
| `Mercado de Aviação` ⓐ | Aviation market | aggregate total |
| `Gasolina de aviação` | Aviation gasoline (avgas) | |
| `Jet` | Jet fuel (Jet A-1) | |
| `JP1` | Jet propellant JP-1 | |
| `JP8` | Jet propellant JP-8 | military jet fuel |

### Biofuels
| Column | English | Notes |
|---|---|---|
| `Biocombustíveis` ⓐ | Biofuels | aggregate total |
| `Biodiesel` | Biodiesel (generic) | |
| `Biodiesel incorporado no gasóleo` | Biodiesel blended into diesel | |
| `Biodiesel fornecido diretamente ao mercado` | Biodiesel supplied directly to market | |
| `Biodiesel B10` / `Biodiesel B15` / `Biodiesel B20` | Biodiesel blends B10 / B15 / B20 | |
| `Bioetanol incorporado na gasolina` | Bioethanol blended into gasoline | |
| `Biogasolina incorporada` | Biogasoline incorporated | |

## Naming assumptions (from `data/processed/notas.txt`)

Product names changed over time; for consistency the source workbook assumes:

- **Fuel (EDP + Tejo)** → *Fuelóleo para produção de energia eléctrica*
- **Thin-Fuel-Oil** → *Fuelóleo nº3 BTE*
- **Thick-Fuel-Oil 1%S** → *Fuelóleo nº4 BTE*
- **Thick-Fuel-Oil 3,5%S** → *Fuelóleo nº4 ATE*

Zero observations for **Propileno** were added for 2012–2015.

## Source

DGEG — Direção-Geral de Energia e Geologia (Portugal), monthly oil-product sales
statistics (PMA / PGP series). Eurostat `nrg_*` extracts under `data/raw/eurostat/`
are auxiliary and not part of the wide/time-series outputs.
