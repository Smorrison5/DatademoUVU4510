# DatademoUVU4510

## Journal Entry Analysis Workflow

This repo includes a lightweight analysis workflow for `je_samples.xlsx` that produces summary outputs in the `outputs/` folder. The GitHub Actions workflow (`JE Sample Analysis`) runs the script on demand or whenever the input file or analysis script changes, and uploads the outputs as a downloadable artifact.

### How it works

- **Script:** `scripts/analyze_je_samples.py`
- **Benford Script:** `scripts/benford_analysis.py`
- **Outputs:**
  - `outputs/summary.md` – human-readable overview
  - `outputs/summary.json` – machine-readable summary
  - `outputs/numeric_summary.csv` – descriptive statistics for numeric columns
  - `outputs/benford_summary.md` – Benford's Law overview
  - `outputs/benford_summary.json` – Benford counts and percentages
  - `outputs/benford_summary.csv` – Benford counts and percentages in CSV
  - `outputs/benford_chart.svg` – Benford chart for quick visual comparison

### Run locally

```bash
python scripts/analyze_je_samples.py
```

To run the Benford analysis against the first numeric column with at least 10 values:

```bash
python scripts/benford_analysis.py
```

To target a specific column and file:

```bash
python scripts/benford_analysis.py --file your_file.xlsx --column "Amount"
```

### Run in GitHub Actions

1. Go to the **Actions** tab.
2. Select **JE Sample Analysis**.
3. Click **Run workflow**.
4. Download the `je-sample-outputs` artifact from the completed run.
