## Less memory usage - more speed

A Python package to reduce the memory usage of pandas DataFrames without changing the underlying data. It speeds up your workflow and reduces the risk of running out of memory.

## Installation

```python
pip install a-pandas-ex-less-memory-more-speed
```

```python
from a_pandas_ex_less_memory_more_speed import pd_add_less_memory_more_speed
pd_add_less_memory_more_speed()
import pandas as pd
df = pd.read_csv(    "https://github.com/pandas-dev/pandas/raw/main/doc/data/titanic.csv",)
df.ds_reduce_memory_size()

```

## Update 2023/05/04
```python

# to carefully handle callables, iterables and other objects in cells 

df.ds_reduce_memory_size_carefully()


    Optimizes the memory usage of a pandas DataFrame or Series by converting data types and reducing memory size.

    Args:
    df_ (pd.Series | pd.DataFrame): The DataFrame or Series to be optimized.
    ignore_columns (tuple | list, optional): A tuple or list of column names to ignore during optimization. Defaults to ().
    not_allowed_to_convert (tuple | list, optional): A tuple or list of modules that should not be converted during optimization. Defaults to ("shapely",).
    allowed_to_convert (tuple | list, optional): A tuple or list of modules that are allowed to be converted during optimization. Defaults to ("pandas", "numpy").
    include_empty_iters_in_pd_na (bool, optional): If True, empty iterators will be converted to pd.NA during optimization. Defaults to False.
    include_0_len_string_in_pd_na (bool, optional): If True, zero-length strings will be converted to pd.NA during optimization. Defaults to False.
    verbose (bool, optional): If True, print information about the memory usage before and after optimization. Defaults to True.

    Returns:
    pd.DataFrame | pd.Series: The optimized DataFrame or Series.

    Raises:
    None.
    
    
```

## Update 2022/10/08

```python
#added pandas.Series.ds_optimize_int / pandas.DataFrame.ds_optimize_int
#to optimize only ints

     PassengerId  Survived  Pclass  ...     Fare Cabin  Embarked
0              1         0       3  ...   7.2500   NaN         S
1              2         1       1  ...  71.2833   C85         C
2              3         1       3  ...   7.9250   NaN         S
3              4         1       1  ...  53.1000  C123         S
4              5         0       3  ...   8.0500   NaN         S
..           ...       ...     ...  ...      ...   ...       ...
886          887         0       2  ...  13.0000   NaN         S
887          888         1       1  ...  30.0000   B42         S
888          889         0       3  ...  23.4500   NaN         S
889          890         1       1  ...  30.0000  C148         C
890          891         0       3  ...   7.7500   NaN         Q
[891 rows x 12 columns]    


df.ds_optimize_int()
df.PassengerId: Using dtype: np.uint16
df.Survived: Using dtype: np.uint8
df.Pclass: Using dtype: np.uint8
df.SibSp: Using dtype: np.uint8
df.Parch: Using dtype: np.uint8
Out[7]: 
     PassengerId  Survived  Pclass  ...     Fare Cabin  Embarked
0              1         0       3  ...   7.2500   NaN         S
1              2         1       1  ...  71.2833   C85         C
2              3         1       3  ...   7.9250   NaN         S
3              4         1       1  ...  53.1000  C123         S
4              5         0       3  ...   8.0500   NaN         S
..           ...       ...     ...  ...      ...   ...       ...
886          887         0       2  ...  13.0000   NaN         S
887          888         1       1  ...  30.0000   B42         S
888          889         0       3  ...  23.4500   NaN         S
889          890         1       1  ...  30.0000  C148         C
890          891         0       3  ...   7.7500   NaN         Q
```

## Usage

```python
df = pd.read_csv(    "https://github.com/pandas-dev/pandas/raw/main/doc/data/titanic.csv",)
from random import choice

#Let's add some more data types
truefalse = lambda: choice([True, False])
df['truefalse'] = [truefalse() for x in range(len(df))]

df['onlynan'] = pd.NA

df['nestedlists'] = [[[1]*10]] * len(df)

mixedstuff = lambda: choice([True, False, 'right', 'wrong', 1,2,23,343.555,23.444, [442,553,44], [],''])
df['mixedstuff'] =[mixedstuff() for x in range(len(df))]

floatnumbers = lambda: choice([33.44,344.42424265,15.0,3222.33])
df['floatnumbers']=[floatnumbers() for x in range(len(df))]

floatnumbers0 = lambda: choice([33.0,344.0,15.0,3222.0])
df['floatnumbers0']=[floatnumbers0() for x in range(len(df))]

intwithnan = lambda: choice([1,2,3,4,5,pd.NA])
df['intwithnan']=[intwithnan() for x in range(len(df))]


df2 = optimize_dtypes(
    dframe=df,
    point_zero_to_int=True,
    categorylimit=15,
    verbose=True,
    include_na_strings_in_pd_na=True,
    include_empty_iters_in_pd_na=True,
    include_0_len_string_in_pd_na=True,
    convert_float=True,
    check_float_difference=True,
    float_tolerance_negative=-0.1,
    float_tolerance_positive=0.1,
)
print(df)
print(df2)
print(df.dtypes)
print(df2.dtypes)


Memory usage of dataframe is: 0.12333202362060547 MB
█████████████████████████████
Analyzing df.PassengerId
----------------
df.PassengerId Is numeric!
df.PassengerId Max: 891
df.PassengerId Min: 1
df.PassengerId: Only .000 in columns -> Using int - Checking which size fits best ...
df.PassengerId: Using dtype: np.uint16
█████████████████████████████
Analyzing df.Survived
----------------
df.Survived Is numeric!
df.Survived Max: 1
df.Survived Min: 0
df.Survived: Only .000 in columns -> Using int - Checking which size fits best ...
df.Survived: Using dtype: np.uint8
█████████████████████████████
Analyzing df.Pclass
----------------
df.Pclass Is numeric!
df.Pclass Max: 3
df.Pclass Min: 1
df.Pclass: Only .000 in columns -> Using int - Checking which size fits best ...
df.Pclass: Using dtype: np.uint8
█████████████████████████████
Analyzing df.Name
----------------
df.Name: Using dtype: string
█████████████████████████████
Analyzing df.Sex
----------------
df.Sex: Using dtype: category
█████████████████████████████
Analyzing df.Age
----------------
df.Age Is numeric!
df.Age Max: 80.0
df.Age Min: 0.42
df.Age: Using dtype: Float64
█████████████████████████████
Analyzing df.SibSp
----------------
df.SibSp Is numeric!
df.SibSp Max: 8
df.SibSp Min: 0
df.SibSp: Only .000 in columns -> Using int - Checking which size fits best ...
df.SibSp: Using dtype: np.uint8
█████████████████████████████
Analyzing df.Parch
----------------
df.Parch Is numeric!
df.Parch Max: 6
df.Parch Min: 0
df.Parch: Only .000 in columns -> Using int - Checking which size fits best ...
df.Parch: Using dtype: np.uint8
█████████████████████████████
Analyzing df.Ticket
----------------
df.Ticket: Using dtype: string
█████████████████████████████
Analyzing df.Fare
----------------
df.Fare Is numeric!
df.Fare Max: 512.3292
df.Fare Min: 0.0
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Max. positive difference - limit 0.1
498   -0.05
305   -0.05
708   -0.05
Max. negative difference - limit -0.1
679    0.1708
258    0.1708
737    0.1708
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
------------- <class 'numpy.float16'> ------------- not right for df.Fare
Checking next dtype...
True -> within the desired range: 0.1 / -0.1
False      5
True     886
-------------------
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Max. positive difference - limit 0.1
0      0.0
587    0.0
588    0.0
Max. negative difference - limit -0.1
0      0.0
598    0.0
587    0.0
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
+++++++++++++ <class 'numpy.float32'> +++++++++++++ right for df.Fare
True -> within the desired range: 0.1 / -0.1
True    891
-------------------
df.Fare: Using dtype: np.float32
█████████████████████████████
Analyzing df.Cabin
----------------
df.Cabin: Using dtype: string
█████████████████████████████
Analyzing df.Embarked
----------------
df.Embarked: Using dtype: category
█████████████████████████████
Analyzing df.truefalse
----------------
df.truefalse: Using dtype: np.bool_
█████████████████████████████
Analyzing df.onlynan
----------------
df.onlynan Is numeric!
df.onlynan Max: nan
df.onlynan Min: nan
df.onlynan: Only nan in column, continue ...
█████████████████████████████
Analyzing df.nestedlists
----------------
█████████████████████████████
Analyzing df.mixedstuff
----------------
█████████████████████████████
Analyzing df.floatnumbers
----------------
df.floatnumbers Is numeric!
df.floatnumbers Max: 3222.33
df.floatnumbers Min: 15.0
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Max. positive difference - limit 0.1
890   -0.33
597   -0.33
592   -0.33
Max. negative difference - limit -0.1
527    0.075757
190    0.075757
171    0.075757
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
------------- <class 'numpy.float16'> ------------- not right for df.floatnumbers
Checking next dtype...
True -> within the desired range: 0.1 / -0.1
False    219
True     672
-------------------
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Max. positive difference - limit 0.1
0      0.0
587    0.0
588    0.0
Max. negative difference - limit -0.1
0      0.0
598    0.0
587    0.0
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
+++++++++++++ <class 'numpy.float32'> +++++++++++++ right for df.floatnumbers
True -> within the desired range: 0.1 / -0.1
True    891
-------------------
df.floatnumbers: Using dtype: np.float32
█████████████████████████████
Analyzing df.floatnumbers0
----------------
df.floatnumbers0 Is numeric!
df.floatnumbers0 Max: 3222.0
df.floatnumbers0 Min: 15.0
df.floatnumbers0: Only .000 in columns -> Using int - Checking which size fits best ...
df.floatnumbers0: Using dtype: np.uint16
█████████████████████████████
Analyzing df.intwithnan
----------------
df.intwithnan Is numeric!
df.intwithnan Max: 5
df.intwithnan Min: 1
df.intwithnan: Only .000 in columns -> Using int - Checking which size fits best ...
df.intwithnan: Using dtype: Int64
█████████████████████████████
Memory usage of dataframe was: 0.12333202362060547 MB
Memory usage of dataframe is now: 0.07259273529052734 MB
This is  58.85959960718511 % of the initial size
█████████████████████████████
█████████████████████████████
     PassengerId  Survived  Pclass  ... floatnumbers floatnumbers0  intwithnan
0              1         0       3  ...    33.440000          33.0           4
1              2         1       1  ...  3222.330000          15.0           5
2              3         1       3  ...    33.440000          33.0           3
3              4         1       1  ...    15.000000          33.0           1
4              5         0       3  ...    15.000000         344.0           2
..           ...       ...     ...  ...          ...           ...         ...
886          887         0       2  ...   344.424243         344.0           5
887          888         1       1  ...    15.000000          15.0           4
888          889         0       3  ...   344.424243        3222.0           2
889          890         1       1  ...   344.424243        3222.0           4
890          891         0       3  ...  3222.330000        3222.0        <NA>
[891 rows x 19 columns]
     PassengerId  Survived  Pclass  ... floatnumbers floatnumbers0  intwithnan
0              1         0       3  ...    33.439999            33           4
1              2         1       1  ...  3222.330078            15           5
2              3         1       3  ...    33.439999            33           3
3              4         1       1  ...    15.000000            33           1
4              5         0       3  ...    15.000000           344           2
..           ...       ...     ...  ...          ...           ...         ...
886          887         0       2  ...   344.424255           344           5
887          888         1       1  ...    15.000000            15           4
888          889         0       3  ...   344.424255          3222           2
889          890         1       1  ...   344.424255          3222           4
890          891         0       3  ...  3222.330078          3222        <NA>
[891 rows x 19 columns]
PassengerId        int64
Survived           int64
Pclass             int64
Name              object
Sex               object
Age              float64
SibSp              int64
Parch              int64
Ticket            object
Fare             float64
Cabin             object
Embarked          object
truefalse           bool
onlynan           object
nestedlists       object
mixedstuff        object
floatnumbers     float64
floatnumbers0    float64
intwithnan        object
dtype: object
PassengerId        uint16
Survived            uint8
Pclass              uint8
Name               string
Sex              category
Age               Float64
SibSp               uint8
Parch               uint8
Ticket             string
Fare              float32
Cabin              string
Embarked         category
truefalse            bool
onlynan            object
nestedlists        object
mixedstuff         object
floatnumbers      float32
floatnumbers0      uint16
intwithnan          Int64
dtype: object

    Parameters:
        dframe: Union[pd.Series, pd.DataFrame]
            pd.Series, pd.DataFrame
        point_zero_to_int: bool
            Convert float to int if all float numbers in the column end with .0+
            (default = True)
        categorylimit: int
            Convert strings to category, when ratio len(df) / len(df.value_counts) >= categorylimit
            (default = 4)
        verbose: bool
            Keep track of what is happening
            (default = True)
        include_na_strings_in_pd_na: bool
            When True -> treated as nan:

            [
            "<NA>",
            "<NAN>",
            "<nan>",
            "np.nan",
            "NoneType",
            "None",
            "-1.#IND",
            "1.#QNAN",
            "1.#IND",
            "-1.#QNAN",
            "#N/A N/A",
            "#N/A",
            "N/A",
            "n/a",
            "NA",
            "#NA",
            "NULL",
            "null",
            "NaN",
            "-NaN",
            "nan",
            "-nan",
            ]

            (default =True)
        include_empty_iters_in_pd_na: bool
            When True -> [], {} are treated as nan (default = False )

        include_0_len_string_in_pd_na: bool
            When True -> '' is treated as nan (default = False )
        convert_float: bool
            Don't convert columns containing float numbers.
            Comparing the 2 dataframes from the example, one can see that float numbers frequently
            don't have the exact same value as the original float number.
            If decimal digits are important for your work, disable it!
            (default=True)
        check_float_difference: bool
            If a little difference between float dtypes is fine for you, use True
            Ignored if convert_float=False
            (default=True)
        float_tolerance_negative: float

            The negative tolerance you can live with, e.g.
            3222.330078 - 3222.330000 = 0.000078 is fine for you

            Ignored if convert_float=False
            (default= 0)

        float_tolerance_positive: float = 0,
            The positive tolerance you can live with
            3222.340078 - 3222.330000 = 0.010078 is fine for you
             Ignored if convert_float=False
            (default= 0.05)

    Returns:
        Union[pd.DataFrame, pd.Series]
```
