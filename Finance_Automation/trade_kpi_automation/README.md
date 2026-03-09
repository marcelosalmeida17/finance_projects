Trade KPI Automation

Python automation script for daily trading activity KPIs.
Overview

This script processes trading activity files and automatically generates daily KPI reports by calculating the sum of all type of colalteral trades alive in a given set of trading books.

It reads trading and cash flow data, classifies transactions by type (bilateral, triparty, cash), and produces a summarized KPI output.
Features

    Automated daily file detection based on date
    Trade classification using transaction attributes
    KPI aggregation and reporting
    Export of daily metrics to CSV

Technologies

    Python
    Pandas
    File automation with Pathlib

Example Output

KPIs generated:

    NC Bilateral Principal Trades
    NC Bilateral Collateral Trades
    Triparty Principal Trades
    Triparty Collateral Trades
    Cash Principal Trades
    Cash Bookings

Use Case

Demonstrates automation of operational reporting workflows in financial markets environments.
