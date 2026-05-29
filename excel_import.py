import pandas as pd


def load_emails_from_excel(file_path):

    df = pd.read_excel(file_path)

    if df.empty:
        return []

    first_column = df.iloc[:, 0]

    emails = []

    for value in first_column:
        if pd.notna(value):
            emails.append(str(value).strip())

    return emails