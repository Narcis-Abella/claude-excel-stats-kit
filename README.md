# claude-excel-stats-kit

Most researchers don't have a statistics problem. They have a tooling problem.

You know which test to run. You know what the results should look like. But between a raw dataset and a publication-ready Methods section, there are hours of work that have nothing to do with your research: cleaning inconsistent data, navigating R syntax you use twice a year, hunting through SPSS menus, formatting tables, writing the same boilerplate statistical prose you've written a dozen times before.

This kit eliminates that overhead.

Paste one prompt into Claude for Excel and it becomes a full academic statistics agent — one that handles the pipeline from raw data to IEEE-formatted results, so you can spend your time on the science instead of the scaffolding.

---

## Quick start

### Prerequisites
- Microsoft 365 with the [Claude for Excel](https://www.anthropic.com/claude-for-excel) add-in (requires Claude Pro, ~$20/month)
- A free [XLSTAT Cloud](https://app.xlstat.com) account — used for formal hypothesis tests that Excel cannot run natively (Shapiro-Wilk, Welch ANOVA, Kruskal-Wallis, Breusch-Pagan)

### Setup (one time, ~2 minutes)

1. Download `ClaudeForExcel_StatsKit.xlsx` from this repository
2. Open it in Excel Online
3. Open the Claude for Excel sidebar
4. Copy the full text from cell **B2** of sheet `00_MASTER_PROMPT`
5. Paste it into the Claude for Excel chat

Done. Claude will autonomously create all 10 skills and update its persistent instructions — no further input needed.

### Run your first analysis

Once configured, say:

> `"Academic analysis"` or `/academic-stats-protocol`

Claude will ask for your response variable and grouping factors, then run the full pipeline.

---

## What gets installed

### Persistent instructions
Global statistical rules added to Claude for Excel's memory — sensible defaults grounded in peer-reviewed methodology, a clear capability map of what to do in Excel vs. what to delegate to XLSTAT Cloud, and IEEE reporting format rules.

### 10 skills

| Skill | Purpose |
|-------|---------|
| `academic-stats-protocol` | Master orchestrator — runs the full pipeline |
| `stats-descriptive` | EDA: mean, SD, CI, skewness, kurtosis, automatic flags |
| `stats-normality` | Shapiro-Wilk via XLSTAT + distribution fitting |
| `stats-homogeneity` | Levene / Bartlett via XLSTAT |
| `stats-anova` | Welch ANOVA / Kruskal-Wallis via XLSTAT |
| `stats-posthoc` | Games-Howell / Tukey HSD / Dunn + Holm-Bonferroni |
| `stats-regression` | OLS via LINEST(), VIF, model comparison by AIC/BIC |
| `stats-residuals` | Cook's D, Q-Q, Breusch-Pagan, Durbin-Watson |
| `stats-report` | IEEE Methods + Results text generation |
| `xlstat-guide` | Step-by-step XLSTAT Cloud navigation |

---

## The pipeline

| Phase | What happens | Where |
|-------|-------------|-------|
| 0 — Intake | Variable identification, outlier detection, analysis plan | Excel |
| 1 — EDA | Descriptive stats with auditable formulas, automatic flags | Excel |
| 2 — Normality | Shapiro-Wilk + Anderson-Darling, parametric/non-parametric decision | XLSTAT Cloud |
| 3 — Homoscedasticity | Levene's test (informative, does not change test choice) | XLSTAT Cloud |
| 4 — Inference | Welch ANOVA or Kruskal-Wallis | XLSTAT Cloud |
| 5 — Post-hoc | Games-Howell / Dunn + Holm-Bonferroni, effect sizes, letter groups | XLSTAT Cloud |
| 6 — Regression | OLS with LINEST(), model comparison by AIC/BIC | Excel |
| 7 — Diagnostics | Residuals, Cook's D, Breusch-Pagan, Durbin-Watson | Excel + XLSTAT |
| 8 — Report | IEEE Methods + Results sections, copy-paste ready | Excel |

---

## Statistical methodology

Defaults follow peer-reviewed recommendations, not software convenience:

| Decision | Default | Reference |
|----------|---------|-----------|
| Normality test | Shapiro-Wilk (K-S only if n > 2000) | Razali & Wah (2011) |
| ANOVA | Welch unconditionally — robust to unequal variances without power loss | Delacre et al. (2019) |
| Post-hoc (Welch) | Games-Howell | Games & Howell (1976) |
| Post-hoc (KW) | Dunn + Holm-Bonferroni | Holm (1979) |
| Effect size (ANOVA) | Partial η² = SS_factor / (SS_factor + SS_error) | Cohen (1973) |
| Effect size (KW) | η²_H = (H − k + 1) / (n − k) | Tomczak & Tomczak (2014) |
| Confidence intervals | t-critical via T.INV.2T() — never fixed z = 1.96 | — |
| VIF thresholds | > 5 moderate, > 10 severe | Hair et al. (2010) |

The kit distinguishes clearly between what Claude can do directly in Excel and what requires XLSTAT Cloud. Both paths are covered with step-by-step instructions.

---

## Scope

The 10 skills define a **minimum guaranteed pipeline**, not a ceiling. If your analysis requires methods outside the pipeline — repeated measures ANOVA, MANOVA, mixed models, PCA, survival analysis — Claude proceeds using its general statistical knowledge and flags when it is going beyond the structured pipeline.

---

## Roadmap

This kit covers the core statistical pipeline for between-subjects designs. Planned additions:

- **Data cleaning skills** — audit and cleaning pipeline for messy datasets before they enter the analysis
- **Domain-specific skills** — repeated measures, survival analysis, PCA, mixed models
- **Expanded XLSTAT coverage** — dedicated skills for specific XLSTAT Cloud workflows

The roadmap is driven by real use cases. If you have a research context where this kit falls short, open an issue — that's the most useful contribution you can make.

---

## Contributing

This project grows through community input. There are several ways to contribute:

**Report problems.** If a skill produces a methodologically incorrect result, wrong formula, or bad XLSTAT instructions — open an issue with as much detail as you can. Statistical errors are the highest priority.

**Suggest or submit new skills.** If you work in a research area with specific analytical needs not covered here (clinical trials, longitudinal data, genomics, social sciences...), open an issue describing the use case. If you want to build the skill yourself, submit a pull request — each skill follows a simple 6-field structure (name, trigger, inputs, outputs, restrictions, instructions) that is easy to replicate from the existing examples.

**Share real-world feedback.** If you've used this kit on an actual dataset, feedback on what worked and what didn't is extremely valuable — even if it's just a comment in an issue.

The goal is a community-maintained library of statistical skills for academic researchers, not a single-author project.

---

## Limitations

- Requires Claude Pro (~$20/month) and Microsoft 365 for the Claude for Excel add-in
- XLSTAT Cloud free version covers all tests in this pipeline; some advanced modules require the paid tier
- Designed and tested for between-subjects designs; repeated measures require additional handling
- Claude for Excel's skill and memory system may change — this kit will be updated accordingly

---

## Repository structure

```
claude-excel-stats-kit/
├── ClaudeForExcel_StatsKit.xlsx    # Configuration file — download and use directly
├── README.md
└── LICENSE
```

---

## License

MIT — see [LICENSE](LICENSE).

---

## References

- Akaike, H. (1974). *IEEE Transactions on Automatic Control*, 19(6), 716–723.
- Anderson, T. W., & Darling, D. A. (1954). *Journal of the American Statistical Association*, 49(268), 765–769.
- Breusch, T. S., & Pagan, A. R. (1979). *Econometrica*, 47(5), 1287–1294.
- Cohen, J. (1973). *Educational and Psychological Measurement*, 33(1), 107–112.
- Cohen, J. (1988). *Statistical power analysis for the behavioral sciences* (2nd ed.). Lawrence Erlbaum.
- Cook, R. D. (1977). *Technometrics*, 19(1), 15–18.
- Delacre, M., Leys, C., Mora, Y. L., & Lakens, D. (2019). *International Review of Social Psychology*, 32(1).
- Durbin, J., & Watson, G. S. (1951). *Biometrika*, 38(1–2), 159–177.
- Games, P. A., & Howell, J. F. (1976). *Journal of Educational Statistics*, 1(2), 113–125.
- Hair, J. F., Black, W. C., Babin, B. J., & Anderson, R. E. (2010). *Multivariate data analysis* (7th ed.). Pearson.
- Holm, S. (1979). *Scandinavian Journal of Statistics*, 6(2), 65–70.
- Razali, N. M., & Wah, Y. B. (2011). *Journal of Statistical Modeling and Analytics*, 2(1), 21–33.
- Schwarz, G. (1978). *Annals of Statistics*, 6(2), 461–464.
- Tomczak, M., & Tomczak, E. (2014). *Trends in Sport Sciences*, 1(21), 19–25.