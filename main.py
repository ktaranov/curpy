import os
import numpy as np
import pandas as pd
import warnings

DATA_PATH = "data"


def main():
    """
    Создаёт 2 листа в эксель файле:
    один с пробелами ( выходные и праздники),
    другой без ( последнее известное значение)
    """
    file_paths = [os.path.join(DATA_PATH, i) for i in os.listdir(DATA_PATH)]
    unfilled = pd.DataFrame()
    for path in file_paths:
        # skip styles warning
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            df = pd.read_excel(path, index_col=[0])

        df = df.iloc[:-2]
        unfilled = pd.concat([unfilled, df])

    unfilled = unfilled.replace("---", np.nan)
    unfilled.index = pd.to_datetime(unfilled.index, format="%d %b %Y")
    idx = pd.date_range(unfilled.index.min(), unfilled.index.max(), freq="D")

    unfilled = unfilled.reindex(idx, fill_value=np.nan)
    total = unfilled.ffill()

    with pd.ExcelWriter("Central_Parity_Historical_Data_2006_2023_sorted_and_filled.xlsx", engine="xlsxwriter") as writer:
        unfilled.to_excel(writer, sheet_name='Без праздников')
        total.to_excel(writer, sheet_name='С праздниками и выходными')


if __name__ == "__main__":
    main()