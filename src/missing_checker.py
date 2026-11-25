
def find_missing_values(df):

    # Identify missing values
    missing_mask = df.isnull()

    # Count missing values in each column
    missing_count = missing_mask.sum()

    # Get rows containing missing values
    rows_with_missing = df[missing_mask.any(axis=1)]

    # Calculate total missing values
    total_missing = missing_count.sum()

    return missing_count, rows_with_missing, total_missing
