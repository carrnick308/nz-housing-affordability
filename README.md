# NZ housing affordability calculator

The interactive calculator behind Your Money Blueprint's housing affordability page. It compares median house prices with median household incomes in 10 New Zealand regions (plus the national figure) from 2005: house price to income, repayments as a share of income, and the deposit in years of income.

It's published with GitHub Pages and shown on yourmoneyblueprint.co.nz through a Squarespace Code Block (`embed/squarespace-housing-affordability.html`).

## The code in this repo is scrambled

`assets/housing.js` is deliberately scrambled so the workings can't easily be copied. The readable version is kept privately, outside this repo. Don't replace it with an unscrambled copy.

## Yearly update

1. Add the new year's figures to `inbox/housing-inputs.xlsx` (one new column on each sheet).
2. Upload it to `inbox/`, replacing the old file.
3. The "Update housing data" action rebuilds `data/housing-data.json`. Interest rates come from the nz-mortgage-rates project automatically.

If anything in the file looks wrong, the action stops and the published data is left as it was. The reason is shown in the action's log.

## Data sources

- House prices: REINZ median sale price for August (figures before 2017 from interest.co.nz's REINZ series).
- Incomes: Stats NZ Household Labour Force Survey, median weekly household income, year ended June, times 52.
- Interest rates: Reserve Bank B21 special (from 2017) and B20 standard 1-year fixed rates for August.
