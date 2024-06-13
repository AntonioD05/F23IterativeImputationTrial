from sklearn.experimental import enable_iterative_imputer
from sklearn.impute import IterativeImputer
from sklearn.tree import DecisionTreeRegressor
import pandas as pd
import numpy as np

# Load all sheets into a dictionary of DataFrames
path = 'C:\\Users\\adiaz\\Downloads\\F23_rawSleepData (1).xlsx'
sheets = pd.read_excel(path, sheet_name=None)

# Define the imputer outside the loop to use the same settings for all sheets
imputer = IterativeImputer(estimator=DecisionTreeRegressor(), max_iter=15, random_state=0, tol=0.1)

# Process each sheet
for sheet_name, df in sheets.items():
    # Drop the first column if it's not needed or an index
    df = df.drop(df.columns[0], axis=1)

    # Replace zeros with NaN if they represent missing data
    threshold = 0.50
    for column in df.columns:
        zero_count = (df[column] == 0).sum()
        if zero_count / len(df) <= threshold:
            df[column] = df[column].replace(0, np.nan)

    # Impute missing values using the decision tree regressor
    sheets[sheet_name] = pd.DataFrame(imputer.fit_transform(df), columns=df.columns)

# Save the processed DataFrames back into an Excel file, each sheet separately
output_path = 'C:\\Users\\adiaz\\Downloads\\F23_rawSleepDataCleanedUsingIterativeImputer.xlsx'
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name + 'Cleaned', index=False)

print(f"Data saved to: {output_path}")
