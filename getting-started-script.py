import os

import pandas as pd


def main():
    filename = "first.xlsx"
    file_dir = "myfiles"

    file_path = os.path.join(file_dir, filename)

    df = pd.DataFrame([["ABC", "XYZ"]], columns=["Foo", "Bar"])
    with pd.ExcelWriter(file_path) as writer:
        df.to_excel(writer)

    from datetime import date, datetime

    df = pd.DataFrame(
        [
            [date(2014, 1, 31), date(1999, 9, 24)],
            [datetime(1998, 5, 26, 23, 33, 4), datetime(2014, 2, 28, 13, 5, 13)],
        ],
        index=["Date", "Datetime"],
        columns=["X", "Y"],
    )

    with pd.ExcelWriter(file_path, mode="a", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="dates_test")

    df = pd.read_excel(file_path, index_col=0, sheet_name="Sheet1")
    df = pd.read_excel(file_path, index_col=0, sheet_name="dates_test")
    df = pd.read_excel(file_path, index_col=0, sheet_name="Sheet1")

    tmp1 = df + df

    tmp = "".join([df[col][0] for col in df.columns])
    data = [tmp, tmp[::-1]]
    tmp2 = pd.DataFrame([data], columns=["Foo", "Bar"])

    data = [i.lower() for i in df.iloc[0]]
    tmp3 = pd.DataFrame([data], columns=["Foo", "Bar"])
    # https://pandas.pydata.org/pandas-docs/stable/user_guide/merging.html

    frames = [df, tmp1, tmp2, tmp3]

    result = pd.concat(frames)

    with pd.ExcelWriter(file_path, mode="w") as writer:
        result.to_excel(writer, sheet_name="Sheet1")

    print(pd.read_excel(file_path, index_col=0, sheet_name="Sheet1"))

    os.remove(file_path)


if __name__ == "__main__":
    main()
