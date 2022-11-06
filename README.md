# xlsx-is-not-a-database

a little code to help Tony survive grad school

![pic1](https://github.com/will-wright-eng/xlsx-is-not-a-database/blob/images/pic1.png)

## system requirements

`terminal`

```bash
uname -v
# Darwin Kernel Version 21.6.0: Wed Aug 10 14:25:27 PDT 2022; root:xnu-8020.141.5~2/RELEASE_X86_64

python3 --version
# Python 3.9.13
```

### note

> ideally you would be using jupyter inside of a virtual environment so that there is less risk with dependencies drifting over time

```bash
conda create -n jl python=3.10
conda activate jl

python -m pip install jupyterlab
pip freeze > requirements.txt
```

![pic2](https://github.com/will-wright-eng/xlsx-is-not-a-database/blob/images/pic2.png)

## excel as a database :(

**Bad News:** there are lots of libraries that have attempted to make it easy for python to interact with Excel files
**Good News:** pandas is amazing and has created a solver that combines these libraries

- [xlwt] for xls files
- [xlsxwriter] for xlsx files if xlsxwriter is installed otherwise [openpyxl]
- [odswriter] for ods files

This way you only need to understand how to use [pandas.ExcelWriter]

> I'm going to assume you are using `.xlsx` files

[xlwt]: https://pypi.org/project/xlwt/
[xlsxwriter]: https://pypi.org/project/XlsxWriter/
[odswriter]: https://pypi.org/project/odswriter/
[openpyxl]: https://pypi.org/project/openpyxl/

[pandas.ExcelWriter]: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html

### Note

if your only goal is *only* to write data to an Excel file then [pandas.DataFrame.to_excel] is an easy to use option. the following dives into ExcelWriter because it gives you greater flexibility around writing and appending data to existing files

[pandas.DataFrame.to_excel]: https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html

## getting started

### setup

1. create directory & files
2. install pandas
3. start a notebook

```bash
mkdir myfiles
touch myfiles/test-data.xlsx

# the "3" in python3 assumes you aren't using a virtual environment
python3 -m pip install pandas
python3 -m pip install XlsxWriter
python3 -m pip install openpyxl

jupyter lab
```

### CRUD

*see `getting-started.ipynb` so that you can execute each step*

1. create
2. read
3. update
4. delete

## example code

*here is some code I use in production*

```python
def export_to_excel(outputs: dict, export_file_name: str, REPORT_FILE_PATH: str):
    writer = pd.ExcelWriter(f"{REPORT_FILE_PATH}/{export_file_name}.xlsx")
    if "table_of_contents" in list(outputs):
        df = outputs["table_of_contents"]
        df.to_excel(writer, sheet_name="table_of_contents")
        outputs.pop("table_of_contents")

    for table in outputs:
        df = outputs[table]
        df.to_excel(writer, sheet_name=table)
    writer.save()
    return f"export to {export_file_name}.xlsx complete"
```

- this for loop takes a dictionary input containing data and metadata -- these two operations in particular are useful:
  - `pd.pivot_table`
  - `df.groupby`

```python
for table_name, inputs in input_dict.items():
    print(inputs["type"], f"len df:{str(len(df))}")

    if inputs["type"] == "data_filter":
        df = df.loc[inputs["bool_op"](df[inputs["column"]], inputs["bool_arg"])]

    elif inputs["type"] == "data_reset":
        df = raw_df

    elif inputs["type"] == "pivot_table":
        df.loc[:, inputs["values"]] = df.loc[:, inputs["values"]].astype(float)
        table = pd.pivot_table(
            df,
            values=inputs["values"],
            index=inputs["index"],
            columns=inputs["columns"],
            aggfunc=np.sum,
        )
        table.columns = [j for i, j in list(table.columns)]
        if len(table) > 0:
            outputs[table_name] = table

    elif inputs["type"] == "groupby_table":
        headers = {i: inputs["aggfuncs"] for i in inputs["values"]}
        table = df.groupby(inputs["index"]).agg(headers)
        if len(table) > 0:
            outputs[table_name] = table

    elif inputs["type"] == "sum_on_previous_table":
        tmp = pd.DataFrame(table.sum(axis=inputs["axis"]))
        tmp.columns = ["sum"]
        outputs[table_name] = tmp

    else:
        print("invalid input type")
```
