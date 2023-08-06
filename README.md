## pyfvp

pyvfp is a pyhon library for interacting with Visual Fox Pro databases using the `oledbvfp` driver on 64-bit platforms.

It leverages a c++ 32bit binary to execute and returns queries in text format. Selects are returned with the associated
data types and coerced into valid python types via pandas.

Both the executable and source are included in the package, and the binary file graciously handles errors so no table
locks should be encountered. All queries are run NON exclusively.

### Installation

```bash
pip install pyvfp
```

### Usage

#### Selecting data

To select data, use the select method. This will return a pandas dataframe with the data types coerced to valid python
types. There are also optional kwargs to return the data as a list of dicts, or a list of lists, etc.

```python

from pyvfp import Vfp

vfp = Vfp("C:/path/to/your/dbc/my_dbc.DBC")
df = vfp.select("select * from my_table")

```

You can then save, or manipulate the data to your will.

```python

df.to_csv("my_table.xlsx")
df.to_excel("my_table.xlsx")

```

#### Updating, Inserting and Deleting

To update, delete or insert, use the execute method. This will return a boolean True, if successful for raise and
exception.

```python

from pyvfp import Vfp

vfp = Vfp("C:/path/to/your/dbc/my_dbc.DBC")
vfp.execute("update my_table set my_field = 'my_value' where my_field = 'my_other_value'")

```
