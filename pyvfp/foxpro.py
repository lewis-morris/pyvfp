import subprocess
from importlib import resources
from subprocess import CompletedProcess
from typing import Union

import pandas as pd
from pandas import DataFrame

from pyvfp import bin


def _convert_dtype(dtype: str):
    """
    Convert a dtype to pandas' data type

    :param dtype: String, type of the data
    :return: pandas data type
    """
    if dtype == "float":
        return float
    elif dtype == "int":
        return int
    elif dtype == "bool":
        return bool
    elif dtype == "date":
        return 'datetime64[ns]'
    else:
        return dtype


class Vfp:
    """
    Class to query a Visual FoxPro database (.DBC) using the VFP executable, written in c++.
    The .exe file can also be used in the console.
    """

    def __init__(self, dbc_container: str):
        """
        Initialize Query with an executable path

        :param dbc_container: String, path of the database container (.DBC) you want to query.
        """
        self.dbc_container: str = dbc_container

    def _get_result(self, query: str) -> CompletedProcess[str]:
        """
        Internal method to run the query and return the result
        :param query: text of the query
        :return: str: result of the query
        :raises: RuntimeError
        """
        with resources.as_file(resources.files(bin) / "vfp.exe") as vfp_exe_path:
            result: CompletedProcess[str] = subprocess.run([str(vfp_exe_path)] + [self.dbc_container, query],
                                                           stdout=subprocess.PIPE,
                                                           stderr=subprocess.PIPE,
                                                           text=True)

        if result.returncode != 0:
            raise RuntimeError(f"Query failed: {result.stderr}")

        return result

    def select(self, query: str, **kwargs) -> Union[dict | list | str | DataFrame]:
        """
        Run the query with provided arguments and return the results as a panda's DataFrame or another type if
        specified in the kwargs, only to be used to return results and not to modify the database.

        :param query: The query to run on the database. The query should be a string.
        :param kwargs: If you want the output to be another format, add any of the following;
                       as_dict = True, as_list = True, as_string= True
        :return: Pandas DataFrame (default) | Dictionary | List | String
        """

        result = self._get_result(query=query)
        lines = result.stdout.split("\n")

        columns = [x.strip() for x in lines[0].split("|") if x not in ['', ', ']]
        df_data = {k: [] for k in columns}

        types = []
        for x in lines[1:]:
            line_vals = [y.strip().split(":", 1) for y in x.split("|") if y not in ['', ', ']]
            if types:
                for i, val in enumerate(line_vals):
                    df_data[columns[i]].append(val[1])
            else:
                for i, val in enumerate(line_vals):
                    types.append(_convert_dtype(val[0]))
                    df_data[columns[i]].append(val[1])

        df = pd.DataFrame(df_data)

        for i, column in enumerate(columns):
            df[column] = df[column].astype(types[i])

        if kwargs.get("as_dict"):
            return df.to_dict()
        elif kwargs.get("as_list"):
            return df.to_list(index=False)
        elif kwargs.get("as_string"):
            return df.to_string(index=False)
        return df

    def execute(self, query: str) -> bool:
        """
        Execute a command such as update, insert or delete

        :param query: The query to run on the database.

        :return: Boolean
        """
        result = self._get_result(query=query)
        return result.returncode == 0
