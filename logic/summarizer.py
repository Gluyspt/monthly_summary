import pandas as pd
import numpy as np

class Entry:
    def __init__(self, name, month, quantity, weight):
        self.name = name
        self.month = month
        self.quantity = quantity
        self.weight = weight

def read_and_summarize(df, name_col_index, value_col_index=24):
    entries = []

    for idx, row in df.iterrows():
        name = row.iloc[name_col_index]

        # Skip rows where name is 'Exporter' or 'Importer' (case-insensitive)
        if isinstance(name, str) and name.strip().lower() in ('exporter', 'importer'):
            continue

        date_raw = row.iloc[1]  # B
        value = row.iloc[value_col_index]

        try:
            month = pd.to_datetime(date_raw).to_period('M').strftime('%Y-%m')
        except Exception:
            month = 'Unknown'

        value = pd.to_numeric(value, errors='coerce')
        value = 0 if pd.isna(value) else value

        entries.append(Entry(name, month, 0, value))  # quantity unused here for simplicity

    summary = {}
    for e in entries:
        if e.name not in summary:
            summary[e.name] = {}
        if e.month not in summary[e.name]:
            summary[e.name][e.month] = {"total_value": 0}
        summary[e.name][e.month]["total_value"] += e.weight  # using weight to store the value

    names = sorted(summary.keys())
    months = [f"{m:02d}" for m in range(1, 13)]

    data = []
    for name in names:
        row = {"Name": name}
        for m in months:
            total = sum(
                totals["total_value"]
                for ym, totals in summary[name].items()
                if ym.endswith(f"-{m}")
            )
            row[m] = total
        data.append(row)

    df_result = pd.DataFrame(data).set_index("Name")

    # Round up to 3 decimal places
    df_result = np.ceil(df_result * 100) / 100

    return df_result