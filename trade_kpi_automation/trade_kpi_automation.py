from datetime import date, timedelta
import pandas as pd
from pathlib import Path

# Set target date (yesterday) in a portable format
yesterday = date.today() - timedelta(days=1)
target_date = yesterday.strftime("%Y%m%d")

# Place your files under ./data/
# Example filenames:
#   ./data/flows/Flows_20260302_example.csv
flows_folder = Path("./data/flows")
flows_file_path = next(flows_folder.glob(f"Flows_{target_date}*.csv"))

#Filter by user
users = ["USER_A", "USER_B", "USER_C", "USER_D"]

# Import Alive Trades + Flows
alive_trades = pd.read_csv(
    f"./data/alive_trades/Alive_Trades_{target_date}.csv",
    sep=";",
    encoding="utf-8"
)
# Additional File for Cash Trades
flows_file = pd.read_csv(flows_file_path, sep=";", encoding="utf-8")

# Count Non Cash Bilateral Principal Trades
NCBP = (
    alive_trades["Type"].astype(str).str.strip().str.upper().eq("NC_TYPE")
    &
    alive_trades["Book"].astype(str).str.strip().str.upper().str.startswith("NC_BOOK")
    &
    alive_trades["Collateral Profile"].astype(str).str.strip().str.upper().str.startswith("BILATERAL")
)

# Count Non Cash Bilateral Collateral Trades
NCBC = (
    alive_trades["Type"].astype(str).str.strip().str.upper().eq("NC_COLLATERAL_TYPE")
    &
    alive_trades["Book"].astype(str).str.strip().str.upper().eq("NC_BOOK")
    &
    alive_trades["Collateral Profile"].astype(str).str.strip().str.upper().str.startswith("BILATERAL")
)

# Count Ttriparty Principal
TP = (
    alive_trades["Type"].astype(str).str.strip().str.upper().eq("NC_TYPE")
    &
    alive_trades["Collateral Profile"].astype(str).str.strip().str.upper().str.startswith(
        ("TRI_A", "TRI_B", "TRI_C", "TRI_D", "TRI_E")
    )
)

# Count Triparty Collateral
TC = (
    alive_trades["Type"].astype(str).str.strip().str.upper().eq("NC_COLLATERAL_TYPE")
    &
    alive_trades["Book"].astype(str).str.strip().str.upper().eq("NC_TRIPARTY_BOOK")
    &
    alive_trades["Collateral Profile"].astype(str).str.strip().str.upper().str.startswith(
        ("TRI_A", "TRI_B", "TRI_C", "TRI_D", "TRI_E")
    )
)

# Count Cash Principal
CP = (
    alive_trades["Type"].astype(str).str.strip().str.upper().isin(["CASH_TYPE_1", "CASH_TYPE_2"])
)

# Count Cash Bookings
CashBookings = (
    flows_file["INS_TYPE"].astype(str).str.strip().str.upper().isin(["CASH_RECEIVED", "CASH_PAID"])
    &
    flows_file["CREATION_USER"].astype(str).str.strip().str.upper().isin(users)
)

# Output
NCBPcount = NCBP.sum()
NCBCcount = NCBC.sum()
TPcount = TP.sum()
TCcount = TC.sum()
CPcount = CP.sum()
CashBookingsCount = CashBookings.sum()

# Output Dataframe
output_df = pd.DataFrame(
    {
        "Metric": [
            "NC Bilateral Principal Trades",
            "NC Bilateral Collateral Trades",
            "Triparty Principal Trades",
            "Triparty Collateral Trades",
            "Cash Principal Trades",
            "Cash Bookings"
        ],
        "Count": [
            NCBPcount,
            NCBCcount,
            TPcount,
            TCcount,
            CPcount,
            CashBookingsCount
        ]
    }
)

output_path = Path(f"./output/KPIs_Output_{target_date}.csv")
output_path.parent.mkdir(parents=True, exist_ok=True)

output_df.to_csv(output_path, index=False, sep=";", encoding="utf-8-sig")
print(f"Saved: {output_path.resolve()}")
