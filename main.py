import os
import numpy as np
import pandas as pd
import warnings

DATA_PATH = "data"


def main():
    file_paths = [os.path.join(DATA_PATH, i) for i in os.listdir(DATA_PATH)]
    total = pd.DataFrame()
    for path in file_paths:
        # skip styles warning
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            df = pd.read_excel(path, index_col=[0])

        df = df.iloc[:-2]
        total = pd.concat([total, df])

    total = total.replace("---", np.nan)
    total.index = pd.to_datetime(total.index, format="%d %b %Y")
    idx = pd.date_range(total.index.min(), total.index.max(), freq="D")
    total = total.reindex(idx, fill_value=np.nan)
    total = total.ffill()

    total.to_csv("Central_Parity_Historical_Data_2006_2023_sorted_and_filled.csv")


if __name__ == "__main__":
    main()