# Power of One – Member Trends

**Created by BNI Lakefront** (Highland Park, IL) · built by Michael Federman, [Keeping IT Easy](https://keepingiteasy.net) · v1.0 · November 4, 2026

Free for any BNI chapter to use. Not an official BNI Global product; BNI® is a registered trademark of BNI Global, LLC.

Analyze BNI **Power of One** monthly reports per member and see how each person is trending.

**Run it in your browser:** open https://keepingiteasy.github.io/power-of-one/, then drop in two or more monthly
Power of One `.xlsx` files (named starting with the report month, e.g. `2026-10 Power of One.xlsx`).
Files are processed entirely in your browser — nothing is uploaded.

What you get:
- Every member's score by month, 1-month change, change since first report, trend status
  (Improving / Steady / Declining / New) and the category to focus on
- Member drill-down: 11-month score history, category points and raw activity per report
- Chapter and region averages
- Excel downloads: full report, or one chapter at a time

Status rules: Improving = trend ≥ +2 pts/month or ≥ +15 since first report; Declining = ≤ −2 pts/month
or ≤ −15; otherwise Steady (needs 3+ months). Members are matched by name across chapters.

## Offline / batch version
`Scripts/build_reports.py` builds the same analysis as formatted Excel workbooks (master + one per chapter)
from a local `Source Data/` folder. See `Scripts/How to Refresh.txt`.

No member data is stored in this repository.
