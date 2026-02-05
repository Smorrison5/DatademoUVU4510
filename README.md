# DatademoUVU4510

## Journal Entry Analysis Workflow

This repo includes a lightweight analysis workflow for `je_samples.xlsx` that produces summary outputs in the `outputs/` folder. The GitHub Actions workflow (`JE Sample Analysis`) runs the script on demand or whenever the input file or analysis script changes, and uploads the outputs as a downloadable artifact.

### How it works

- **Script:** `scripts/analyze_je_samples.py`
- **Outputs:**
  - `outputs/summary.md` – human-readable overview
  - `outputs/summary.json` – machine-readable summary
  - `outputs/numeric_summary.csv` – descriptive statistics for numeric columns

### Run locally

```bash
python scripts/analyze_je_samples.py
```

### Run in GitHub Actions

1. Go to the **Actions** tab.
2. Select **JE Sample Analysis**.
3. Click **Run workflow**.
4. Download the `je-sample-outputs` artifact from the completed run.
