import pandas as pd
import numpy as np

df = pd.read_excel("D:\Incidents_CC_BA_NA_OCT_2018.xlsx", skiprows=3)

df.drop(df.columns[[0]], axis=1, inplace=True)

# filter out Europe and Asia tickets
df = df[(df.Country != "Europe")]
df = df[(df.Country != "Asia")]

# filter to only keep ones with IW
df = df[df['Revised Resolved Group'].str.contains("IW")]

df=df.reset_index()
df=df.drop(['index'],axis=1)


# after column, header, value
df.insert(2, "Day", df['Open Time'].dt.weekday_name)    # insert newcolumn
df.insert(4, "Calendar Week", df['Open Time'].dt.week)  # insert new column

df.insert(10, "Topic", "")              # insert new column
df.insert(11, "Cause - WC", "")         # insert new column
df.insert(12, "Solution - WC", "")      # insert new column

df.to_excel("D:\Incidents_CC_BA_NA_OCT_2018.xlsx")
