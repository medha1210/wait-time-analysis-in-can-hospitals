import pandas as pd 
import matplotlib.pyplot as plt
import seaborn as sns


## data loading and pre - cleaning

xls = pd.ExcelFile(r"C:\Users\medha\OneDrive\Desktop\healtcare\wait-times-priority-procedures-in-canada-2025-data-tables-en.xlsx")

df = pd.read_excel(xls, sheet_name="Table 1", header=1)

df.columns = df.columns.str.strip()
df.dropna(how = 'all', inplace = True)

print(df.head())
print("printing column names")
print(df.columns.tolist())
print(df)

## data preprocessing and cleaning

# for this analysis, we want to filter out and keep only the wait times which are 50th percentile and 90th percentile
wait_time_data = ["50th Percentile" , "90th Percentile"]
df_wait_time = df[df["Metric"].isin(wait_time_data)].copy()

# cleaning any n/a -> NaN (to treat and  clean any non numerics)
df_wait_time ["Indicator result"] = pd.to_numeric(df_wait_time["Indicator result"],errors="coerce")

# find any missing values and handle them
print(f"Missing value alert:'\n{df_wait_time.isnull().sum()}")

# extra columns exist which are being dropped
df_wait_time = df_wait_time.drop(columns=["Region", "Column1", "Unnamed: 9", "Unnamed: 10", "Unnamed: 11"])

# handling the missing values in Indicator Results by dropping all n/a's
df_wait_time = df_wait_time.dropna(subset=["Indicator result"])

# Cleaning the "Data year" column to fix the FY's and Q3Q4's
df_wait_time["Data year"] = df_wait_time["Data year"].astype(str)
df_wait_time["Year Extracted"] = df_wait_time["Data year"].str.extract(r'(\d{4})')
df_wait_time = df_wait_time[df_wait_time["Year Extracted"].notna()].copy()
df_wait_time["Year Extracted"] = df_wait_time["Year Extracted"].astype(int)

# checking data after the data clean up
print("checking data after clean up")
print(df_wait_time.info())
print(df_wait_time.head())
print(df_wait_time.columns.tolist())

## Exploratory Data Analysis

# Gathering overall statistics of Indicator results
print(df_wait_time["Indicator result"].describe())

# Gathering how nuch information is available for each province
province_data_count = df_wait_time["Province"].value_counts()
print(province_data_count)

# Count of how much data is available per procedure
procedure_data_count = df_wait_time["Indicator"].value_counts()
print(procedure_data_count)

# Average wait time by Province 
av_by_province = df_wait_time.groupby("Province")["Indicator result"].mean().sort_values()
print(av_by_province)

# Average wait time by Procedure 
av_by_procedure = df_wait_time.groupby("Indicator")["Indicator result"].mean().sort_values()
print(av_by_procedure)

# Average wait time by Province and Metric 
avg_wait_by_province_metric = (
    df_wait_time.groupby(["Province", "Metric"])["Indicator result"]
    .mean()
    .unstack()
)
print(avg_wait_by_province_metric)

# Average wait time by Procedure in each province 
avg_wait_proc_prov = (
    df_wait_time.groupby(["Province", "Indicator"])["Indicator result"]
    .mean()
    .reset_index()
    .sort_values(["Province", "Indicator"])
)
print(avg_wait_proc_prov)  

# Average wait time over the years for provinces and procedure
avg_wait_by_year = (
    df_wait_time.groupby("Data year")["Indicator result"]
    .mean()
    .reset_index()
    .sort_values("Data year")
)
print(avg_wait_by_year)

# Average wait time over the years procedure wise
avg_wait_by_year_proc = (
    df_wait_time.groupby(["Year Extracted", "Indicator"])["Indicator result"]
    .mean()
    .reset_index()
    .sort_values(["Year Extracted", "Indicator"])
)
print(avg_wait_by_year_proc) 

# Average wait time over the years province wise
avg_wait_by_year_prov = (
    df_wait_time.groupby(["Year Extracted", "Province"])["Indicator result"]
    .mean()
    .reset_index()
    .sort_values(["Year Extracted", "Province"])
)
print(avg_wait_by_year_prov) 

# Average wait time over the years province and procedure wise 
avg_wait_by_year_prov_proc = (
    df_wait_time.groupby(["Year Extracted", "Province", "Indicator"])["Indicator result"]
    .mean()
    .reset_index()
    .sort_values(["Year Extracted", "Province", "Indicator"])
)

print(avg_wait_by_year_prov_proc)

#   
top5_per_year = (
    df_wait_time.groupby(["Year Extracted", "Indicator"])["Indicator result"]
    .mean()
    .reset_index()
)

# For each year, get the most waited procedure
top1_per_year = (
    df_wait_time.groupby(["Year Extracted", "Indicator"])["Indicator result"]
    .mean()
    .reset_index()
)

# Sort by year ascending and wait time descending
top1_per_year = top1_per_year.sort_values(["Year Extracted", "Indicator result"], ascending=[True, False])

# Pick the top 1 procedure per year
top1_per_year = top1_per_year.groupby("Year Extracted").head(1)

print(top1_per_year)

# Most waited (top 5) procedures over the years
top5_per_year = (
    df_wait_time.groupby(["Year Extracted", "Indicator"])["Indicator result"]
    .mean()
    .reset_index()
)

# For each year, we get top 5 procedures by wait time
top5_per_year = top5_per_year.sort_values(["Year Extracted", "Indicator result"], ascending=[True, False])

# Now we pick top 5 per year
top5_per_year = top5_per_year.groupby("Year Extracted").head(5)

print(top5_per_year)

# Calculating which procedures have high standard of deviations over the years
trend_volatility = (
    df_wait_time.groupby(["Province", "Indicator"])["Indicator result"]
    .agg(['mean', 'std'])
    .reset_index()
)
trend_volatility['coef_var'] = trend_volatility['std'] / trend_volatility['mean']
print(trend_volatility.sort_values('coef_var', ascending=False))

with pd.ExcelWriter("EDA_results.xlsx") as writer:
    av_by_province.to_excel(writer, sheet_name="Avg_By_Province")
    av_by_procedure.to_excel(writer, sheet_name="Avg_By_Procedure")
    avg_wait_by_province_metric.to_excel(writer, sheet_name="Avg_By_Province_Metric")
    avg_wait_proc_prov.to_excel(writer, sheet_name="Avg_Proc_Prov")
    avg_wait_by_year.to_excel(writer, sheet_name="Avg_By_Year")
    avg_wait_by_year_proc.to_excel(writer, sheet_name="Avg_By_Year_Proc")
    avg_wait_by_year_prov.to_excel(writer, sheet_name="Avg_By_Year_Prov")
    avg_wait_by_year_prov_proc.to_excel(writer, sheet_name="Avg_By_Year_Prov_Proc")
    top1_per_year.to_excel(writer, sheet_name="Top1_Per_Year")
    top5_per_year.to_excel(writer, sheet_name="Top5_Per_Year")
    trend_volatility.to_excel(writer, sheet_name="Trend_Volatility")

print("All results saved in a single Excel file: EDA_results.xlsx")


