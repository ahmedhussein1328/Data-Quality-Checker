
def find_duplicates(df):

    # Identify duplicate rows
    duplicate_mask = df.duplicated(keep="first")

    # Get duplicate rows
    duplicate_rows = df[duplicate_mask]

    # Count duplicate rows
    duplicate_count = duplicate_mask.sum()

    return duplicate_rows, duplicate_count
