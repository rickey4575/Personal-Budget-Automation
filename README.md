# Budget Automation

A local Python budgeting tool that imports bank statement CSV/XLSX files, categorizes transactions using keyword rules and AI, maintains a de-duplicated ledger, and updates a single Excel budget workbook.

The project is designed to run locally so your financial data stays on your machine.

## Features

- Import one statement file or a folder of statements
- Normalize common bank-export columns
- Support signed amount columns or separate debit/credit columns
- Categorize transactions using keyword rules
- Optionally use Gemini AI for remaining uncategorized merchants
- Save a de-duplicated `master_ledger.csv`
- Update one living `Budget_Master.xlsx`
- Keep your manually entered budget amounts intact
- Separate real spending from transfers on the dashboard

## Quick Start

1. Install Python 3.10+.

2. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

3. Try the sample data:

   ```powershell
   python budget_tool.py --input sample_data
   ```

4. Open the generated workbook:

   ```text
   Budget_Master.xlsx
   ```

## Using Your Own Statements

Create a local `statements/` folder and put your bank statement CSV/XLSX files inside it:

```text
statements/
  my-bank-statement.csv
```

Then run:

```powershell
python budget_tool.py --input statements
```

On Windows you can also double-click:

```text
run_budget.bat
```

Close `Budget_Master.xlsx` before running the script so Python can save the updated workbook.

## Gemini AI Categorization

AI categorization is optional. The tool works without an API key.

To use Gemini:

1. Create a Gemini API key in Google AI Studio.
2. Save it as a Windows user environment variable named `GEMINI_API_KEY`.
3. Run:

   ```powershell
   python budget_tool.py --input statements --use-ai-categories
   ```

Or double-click:

```text
run_budget_ai.bat
```

AI behavior:

- Existing keyword rules run first.
- Gemini is asked only about remaining uncategorized merchant groups.
- Every AI decision is logged in `ai_category_decisions.csv`.
- Suggestions matching existing categories are applied automatically only when Gemini confidence is at least `0.8`.
- Applied existing-category decisions are appended to `category_rules.csv`, so future runs need less AI.
- Suggested brand-new categories and lower-confidence suggestions are marked as needing approval and are not applied automatically.

Approving AI decisions:

1. Open `ai_category_decisions.csv`.
2. For any row you accept, set `approved` to `True`.
3. Run the normal budget update again, or double-click `run_budget_ai.bat`.

Approved rows are automatically imported into `category_rules.csv` on the next run. This works for both new categories and lower-confidence existing-category suggestions. You can also manually add reliable merchant/category rows directly to `category_rules.csv`.

When an approved AI decision introduces a new spending category, the next run also adds that category to the `Budget` sheet in `Budget_Master.xlsx` with a default monthly budget of `0`. `Salary`, `Other Income`, `Transfers`, and `Uncategorized` are excluded from automatic budget-category creation.

## Important Files

```text
budget_tool.py           Main Python script
Budget_Template.xlsx     Starter workbook used if Budget_Master.xlsx does not exist
category_rules.csv       Keyword-to-category rules
config.json              Bank statement column mapping
run_budget.bat           Windows launcher without AI
run_budget_ai.bat        Windows launcher with Gemini AI categorization
sample_data/             Fake sample statement for testing
```

Generated local files:

```text
Budget_Master.xlsx
master_ledger.csv
ai_category_decisions.csv
```

These generated files are ignored by Git because they may contain private financial data.

## Config

Edit `config.json` if your bank uses different column names.

For a normal signed amount column:

```json
"amount_style": "signed"
```

For separate debit/credit columns:

```json
"amount_style": "debit_credit"
```

Then update the `debit_credit_columns` candidates in `config.json`.

## Dashboard Logic

The dashboard separates:

- Income
- Real spending
- Transfers
- Uncategorized review items
- Pending AI category approvals
- Budget vs actual spending

`Total Spending` excludes `Transfers`. Transfers have their own `Transfer In`, `Transfer Out`, and `Net Transfers` section.

Uncategorized spending is included in `Total Spending` and shown separately in the `Review` section.

The Dashboard `Review` section also shows how many AI decisions are pending approval. Check `ai_category_decisions.csv` for those rows, approve the ones you accept, or manually add a rule to `category_rules.csv`.

## Privacy

Do not commit real statement files, generated ledgers, generated workbooks, API keys, or AI decision logs.

The included `.gitignore` excludes common private artifacts:

```text
statements/
Budget_Master.xlsx
master_ledger.csv
ai_category_decisions.csv
ai_category_suggestions.csv
output/
.env
```
