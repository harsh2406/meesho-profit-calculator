# Meesho Seller Profit Calculator

A Streamlit web app for calculating profit from a Meesho payment Excel workbook.

## What it does

- Reads the **Order Payments** worksheet.
- Finds **Supplier SKU**, **Live Order Status**, and **Final Settlement Amount**.
- Calculates each Delivered order as:

  `Final Settlement Amount - (SKU unit cost × Quantity)`

- Deducts the absolute settlement amount of each Return order from the Delivered profit.
- Ignores RTO orders completely.
- Asks for the unit purchase cost of each Supplier SKU found in Delivered orders.
- Displays SKU-wise profit and lets you download a multi-sheet Excel report.

## Files

- `app.py` — Streamlit application.
- `requirements.txt` — Python packages installed by Streamlit Community Cloud.

## Run on your computer

Python 3.10 or newer is recommended.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Streamlit will show a local address in your terminal, usually `http://localhost:8501`.

## Publish online using GitHub + Streamlit Community Cloud

1. Sign in to GitHub and create a new repository, for example `meesho-profit-calculator`.
2. Upload `app.py`, `requirements.txt`, and `README.md` to the root of that repository.
3. Open [Streamlit Community Cloud](https://share.streamlit.io/) and sign in using GitHub.
4. Select **Create app** and choose your repository, branch (usually `main`), and main file path `app.py`.
5. Click **Deploy**. Streamlit will publish the app at a `*.streamlit.app` web address.
6. Open that web address from your phone, tablet, or computer. You can bookmark it for future use.

When you update the files in GitHub, Community Cloud redeploys the app from the repository.

## Excel format expected

The workbook must contain a worksheet named exactly `Order Payments`. The app detects these required columns even if the capitalization or spacing varies:

- `Supplier SKU`
- `Live Order Status`
- `Final Settlement Amount`

`Quantity` is optional; if it is absent, the app assumes one unit per order. The Meesho workbook format you supplied has the actual headers on the second Excel row, which this app supports.

## Calculation rules

- **Delivered:** settlement amount less SKU unit cost multiplied by quantity.
- **Return:** the absolute value of the Return settlement amount is deducted from profit. This matches your sample workbook, where Return settlement values are negative (for example, `-258`).
- **RTO:** excluded from both settlement and product-cost calculations.
- Other or blank statuses are shown as unrecognized and are not included.

The app does not store the uploaded workbook in a database or write it to permanent app storage. The file is processed in the active Streamlit session.
