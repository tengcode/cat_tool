import pandas as pd
import sas7bdat as sas


def read_data(file_type: str,
              path: str,
              dataset: str,
              package: str = 'pandas',
              coding: str = 'utf-8'
              ) -> pd.DataFrame:
    data = pd.DataFrame()
    file = f"{path}/{dataset}"
    if file_type == 'sas7bdat':
        if package == 'sas7bdat':
            data = sas.SAS7BDAT(f"{file}.sas7bdat").to_data_frame()
        elif package == 'pandas':
            data = pd.read_sas(f"{file}.sas7bdat", encoding=coding)
    elif file_type == 'csv':
        data = pd.read_csv(f"{file}.csv", encoding=coding, engine='pyarrow')
        print(data.columns)
    elif file_type == 'xlsx':
        data = pd.read_excel(f"{file}.xlsx")

    data.columns = [column.lower() for column in data.columns]

    return data

