# Health Expenditure vs. Life Expectancy

**Big Idea:** America's longevity problem is not a budget problem — it is upstream of the budget.

A reproducible Python presentation built on World Bank World Development Indicators, extended with OECD, WHO, and UN data. The deck follows [Storytelling with Data](https://www.storytellingwithdata.com/) principles: Situation → Complication → Resolution.

---

## Slides

| # | Section | Story beat |
|---|---------|------------|
| 1 | Cover | Big Idea stated |
| 2 | The Setup | Global scatter — spending buys longevity until ~$4,000/person |
| 3 | Factor 1 — Education | PISA 2022 scores vs. life expectancy: Japan 533, U.S. 489 |
| 4 | Factor 2 — Lifestyle | Adult obesity rates: U.S. 42.9% vs. Japan 4.9% |
| 5 | Factor 3 — Demographics | Age 65+ share vs. life expectancy — Japan is oldest *and* longest-lived |
| 6 | Factor 4 — Healthcare Costs | Unit price, not utilisation: U.S. prices run 43% above the OECD average |
| 7 | The Pattern | Side-by-side U.S. vs. Japan across all six metrics |
| 8 | The Takeaway | Three recommendations |

---

## Repo layout

```
health-expenditure-analysis/
├── generate_slides.py          # Builds the .pptx from scratch
├── health_expenditure_analysis.pptx
├── requirements.txt
└── data/
    └── health_data.csv         # 58-country dataset, 2000–2023
```

---

## Quickstart

```bash
pip install -r requirements.txt
python generate_slides.py
# Output: health_expenditure_analysis.pptx
```

Requires Python 3.9+.

---

## Data sources

| Dataset | Source | Coverage |
|---------|--------|----------|
| Health expenditure per capita (PPP) | World Bank WDI | 2000–2023 avg; 2022 for cost slide |
| Life expectancy at birth | World Bank WDI | 2000–2023 avg |
| Adult obesity rate (BMI ≥ 30) | WHO Global Health Observatory | 2022 |
| PISA mean score (math · reading · science) | OECD PISA 2022 | 2022 |
| Population aged 65+ (% of total) | UN World Population Prospects / World Bank | 2024 |
| Medical price level (OECD avg = 100) | OECD Society at a Glance 2024 | Latest available |
| Population | World Bank | 2023 |

---

## Design

- **Palette:** dark navy `#1B2A45` for cover/takeaway; white for content slides; red `#C0392B` for U.S. / adverse metrics; teal `#2BA88A` for Japan / positive metrics.
- **Charts:** matplotlib at 150 dpi, embedded as PNG. Gray dots for all countries, coloured highlights for U.S. and Japan.
- **Text:** Mixed-colour headlines via multi-run python-pptx paragraphs.
- Style inspired by the accompanying Power BI analysis (World Bank WDI).
