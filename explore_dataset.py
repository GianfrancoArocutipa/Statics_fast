import pandas as pd
import numpy as np

# Read first few rows of CSV
csv_path = r'c:\Users\PC\Desktop\Estats\archive\Heterogeneous_Data .csv'

print("Loading dataset sample...")
df_sample = pd.read_csv(csv_path, nrows=10)
print("Columns:", df_sample.columns.tolist())

# Now let's see how the 1567 sample was constructed
# In the doc: "1567 repositorios (de 66 641 registros brutos originales)"
# "repositorios de código abierto en Python"
# stars min is 4001, max 19931 (from Table 0 in model doc: Min stars = 4001, Max = 19931, N = 1567)
df_all = pd.read_csv(csv_path)
print(f"Total rows in CSV: {len(df_all)}")
print(f"Languages: {df_all['language'].value_counts().head(10)}")

# Let's filter by language == 'Python'
if 'language' in df_all.columns:
    df_py = df_all[df_all['language'] == 'Python']
    print(f"Python rows: {len(df_py)}")
    print(f"Stars min-max in Python: {df_py['stars'].min()} - {df_py['stars'].max()}")
    
    # Check if there is a stars filter
    df_py_filt = df_py[(df_py['stars'] >= 4000) & (df_py['stars'] <= 20000)]
    print(f"Filtered (4000-20000): {len(df_py_filt)}")
    
    # Check licenses in df_py or df_py_filt
    if 'license' in df_all.columns:
        print("License counts in py filtered:")
        print(df_py_filt['license'].value_counts().head(10))
    elif 'license_spdx_id' in df_all.columns:
        print("License SPDX counts in py filtered:")
        print(df_py_filt['license_spdx_id'].value_counts().head(10))

