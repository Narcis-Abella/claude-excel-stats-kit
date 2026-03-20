# claude-excel-stats-kit

A self-configuration kit that turns **Claude for Excel** into a rigorous academic statistics agent — ready to support peer-reviewed research and IEEE-style reporting.

## What it does

Paste one prompt into Claude for Excel and it autonomously:
- Adds persistent statistical rules to its memory
- Creates 10 specialized skills covering the full academic statistics pipeline

After setup, Claude for Excel can run a complete analysis — from EDA to residual diagnostics — and generate Methods and Results sections in IEEE/APA style, ready to copy-paste into a paper.

---

## Quick start

### 1. Prerequisites
- A Claude for Excel account (Microsoft 365 + Claude extension)
- A free [XLSTAT Cloud](https://app.xlstat.com) account (for formal tests like Shapiro-Wilk, ANOVA, Kruskal-Wallis)

### 2. Generate the Excel configuration file

```bash
pip install openpyxl
python build_stats_kit.py
```

This creates `ClaudeForExcel_StatsKit.xlsx`.

### 3. Self-configure Claude for Excel

1. Open `ClaudeForExcel_StatsKit.xlsx` in Excel Online
2. Open the Claude for Excel sidebar
3. Copy the full text from cell **B2** of sheet `00_MASTER_PROMPT`
4. Paste it into the Claude for Excel chat
5. Claude will autonomously create all 10 skills and update its persistent instructions — no further input needed

### 4. Run your first analysis

Once configured, say:

> "Academic analysis" or "/academic-stats-protocol"

Claude will guide you through the full pipeline with your data.

---

## What gets installed

### Persistent instructions
Global statistical rules added to Claude for Excel's memory:
- Capability map (what to do in Excel vs what to delegate to XLSTAT Cloud)
- Default statistical choices (Welch ANOVA, Holm-Bonferroni, Shapiro-Wilk)
- IEEE reporting format rules
- "Floor not ceiling" policy — skills are a minimum, not a restriction

### 10 Skills

| Sheet | Skill | Purpose |
|-------|-------|---------|
| 02 | `academic-stats-protocol` | Master orchestrator — runs the full pipeline |
| 03 | `stats-descriptive` | EDA: mean, SD, CI, skewness, kurtosis, flags |
| 04 | `stats-normality` | Shapiro-Wilk via XLSTAT + distribution fitting |
| 05 | `stats-homogeneity` | Levene / Bartlett via XLSTAT |
| 06 | `stats-anova` | Welch ANOVA / Kruskal-Wallis via XLSTAT |
| 07 | `stats-posthoc` | Games-Howell / Tukey HSD / Dunn + Holm-Bonferroni |
| 08 | `stats-regression` | OLS via LINEST(), VIF, model comparison |
| 09 | `stats-residuals` | Cook's D, Q-Q, Breusch-Pagan, Durbin-Watson |
| 10 | `stats-report` | IEEE Methods + Results text generation |
| 11 | `xlstat-guide` | Step-by-step XLSTAT Cloud navigation |

---

## Statistical methodology

All decisions follow peer-reviewed recommendations:

| Decision | Rule | Reference |
|----------|------|-----------|
| Normality test | Shapiro-Wilk always (K-S only if n > 2000) | Razali & Wah (2011) |
| Normality verification | Anderson-Darling as secondary check | Anderson & Darling (1954) |
| Default ANOVA | Welch ANOVA unconditionally | Delacre et al. (2019) |
| Post-hoc (Welch) | Games-Howell | Games & Howell (1976) |
| Post-hoc (KW) | Dunn + Holm-Bonferroni | Holm (1979) |
| Effect size (ANOVA) | η²_p = SS_factor / (SS_factor + SS_error) | Cohen (1973) |
| Effect size (KW) | η²_H = (H − k + 1) / (n − k) | Tomczak & Tomczak (2014) |
| Effect size (pairwise) | Cohen's d and r conventions | Cohen (1988) |
| Kurtosis threshold | \|excess kurtosis\| > 2 (Excel KURT() returns excess) | Hair et al. (2010) |
| VIF thresholds | > 5 moderate, > 10 severe | Hair et al. (2010, p. 170) |
| Cook's D flag | 4 / (n − k − 1) | Cook (1977) |
| Heteroscedasticity test | Breusch-Pagan formal test | Breusch & Pagan (1979) |
| Model selection | AIC for model comparison (ΔAIC) | Akaike (1974) |
| Model selection | BIC as complementary criterion | Schwarz (1978) |
| Durbin-Watson | Critical values dL, dU (no fixed cutoffs) | Durbin & Watson (1951) |
| Confidence intervals | T.INV.2T(0.05, n−1) — never fixed z = 1.96 | — |
| p-value interpretation | Never "no effect" without 1−β and effect size | — |
| Multiple comparisons | Holm-Bonferroni > Bonferroni | Holm (1979) |

### Capabilities

**Claude does directly in Excel:**
- Full descriptive statistics with auditable formulas
- OLS regression with LINEST() (Excel 365 spill / Excel 2019 array formula)
- Residuals, standardized residuals, approximate Cook's D, approximate VIF
- Basic diagnostic plots (residuals vs fitted, Q-Q approximation)
- Model comparison tables (AIC, BIC, R²_adj)

**Claude delegates to XLSTAT Cloud (free):**
- Formal normality tests (Shapiro-Wilk, Anderson-Darling)
- Levene / Bartlett homoscedasticity tests
- Welch ANOVA / Factorial ANOVA with Type III SS
- Kruskal-Wallis, Dunn post-hoc
- Breusch-Pagan, Durbin-Watson
- Formal Q-Q plots, violin plots, grouped boxplots

**Beyond the pipeline (floor not ceiling):**
If the analysis requires methods outside the pipeline — Repeated Measures ANOVA, MANOVA, mixed models, robust regression, PCA, survival analysis — Claude proceeds using its general statistical knowledge and suggests JASP or R for models that XLSTAT Cloud cannot handle.

---

## Requirements

- Python 3.7+
- `openpyxl` (`pip install openpyxl`)
- Microsoft 365 with Claude for Excel extension
- Free XLSTAT Cloud account at [app.xlstat.com](https://app.xlstat.com)

---

## Repository structure

```
claude-excel-stats-kit/
├── build_stats_kit.py          # Script that generates the Excel configuration file
├── ClaudeForExcel_StatsKit.xlsx # Pre-built Excel file (ready to use)
├── README.md
└── LICENSE
```

---

## Contributing

Pull requests welcome. If you find a statistical error, methodological inconsistency, or want to add a skill (e.g., survival analysis, PCA, mixed models), please open an issue.

---

## License

MIT — see [LICENSE](LICENSE).

---

## References

- Akaike, H. (1974). A new look at the statistical model identification. *IEEE Transactions on Automatic Control*, 19(6), 716–723.
- Anderson, T. W., & Darling, D. A. (1954). A test of goodness of fit. *Journal of the American Statistical Association*, 49(268), 765–769.
- Breusch, T. S., & Pagan, A. R. (1979). A simple test for heteroscedasticity and random coefficient variation. *Econometrica*, 47(5), 1287–1294.
- Cohen, J. (1973). Eta-squared and partial eta-squared in fixed factor ANOVA designs. *Educational and Psychological Measurement*, 33(1), 107–112.
- Cohen, J. (1988). *Statistical power analysis for the behavioral sciences* (2nd ed.). Lawrence Erlbaum Associates.
- Cook, R. D. (1977). Detection of influential observation in linear regression. *Technometrics*, 19(1), 15–18.
- Delacre, M., Leys, C., Mora, Y. L., & Lakens, D. (2019). Taking parametric assumptions seriously: Arguments for the use of Welch's F-test instead of the classical F-test in one-way ANOVA. *International Review of Social Psychology*, 32(1).
- Durbin, J., & Watson, G. S. (1951). Testing for serial correlation in least squares regression, II. *Biometrika*, 38(1–2), 159–177.
- Games, P. A., & Howell, J. F. (1976). Pairwise multiple comparison procedures with unequal n's and/or variances: A Monte Carlo study. *Journal of Educational Statistics*, 1(2), 113–125.
- Hair, J. F., Black, W. C., Babin, B. J., & Anderson, R. E. (2010). *Multivariate data analysis* (7th ed.). Pearson.
- Holm, S. (1979). A simple sequentially rejective multiple test procedure. *Scandinavian Journal of Statistics*, 6(2), 65–70.
- Razali, N. M., & Wah, Y. B. (2011). Power comparisons of Shapiro-Wilk, Kolmogorov-Smirnov, Lilliefors and Anderson-Darling tests. *Journal of Statistical Modeling and Analytics*, 2(1), 21–33.
- Schwarz, G. (1978). Estimating the dimension of a model. *Annals of Statistics*, 6(2), 461–464.
- Tomczak, M., & Tomczak, E. (2014). The need to report effect size estimates revisited. An overview of some recommended measures of effect size. *Trends in Sport Sciences*, 1(21), 19–25.
