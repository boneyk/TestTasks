import numpy as np
import pandas as pd
from openpyxl import load_workbook
import os

def save_to_excel(df, df_control, df_test, filename="ab_test_data.xlsx"):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="data3")
        df_control.to_excel(writer, index=False, sheet_name="data2")
        df_test.to_excel(writer, index=False, sheet_name="data1")

def tune_test(df):
    df.loc[df['group'] == 'test', 'clicks'] = (df.loc[df['group'] == 'test', 'clicks'] * 1.2).astype(int)
    df.loc[df['group'] == 'test', 'session_length'] = (df.loc[df['group'] == 'test', 'session_length'] * 1.3).astype(int)
    df.loc[df['group'] == 'test', 'purchases'] = (df.loc[df['group'] == 'test', 'purchases'] * 1.5).astype(int)

    # Коррекция количества покупок и кликов
    df['purchases'] = df[['purchases', 'impressions']].min(axis=1)
    df['clicks'] = df[['clicks', 'impressions']].min(axis=1)

    df['CTR'] = df['clicks'] / df['impressions'] * 100
    df['conversion'] = (df['purchases'] / df['impressions']) * 100

    return df

def tune_control(df):
    df.loc[df['group'] == 'control', 'clicks'] = (df.loc[df['group'] == 'control', 'clicks'] * 1.2).astype(int)
    df.loc[df['group'] == 'control', 'session_length'] = (df.loc[df['group'] == 'control', 'session_length'] * 1.3).astype(int)
    df.loc[df['group'] == 'control', 'purchases'] = (df.loc[df['group'] == 'control', 'purchases'] * 1.5).astype(int)

    # Коррекция количества покупок и кликов
    df['purchases'] = df[['purchases', 'impressions']].min(axis=1)
    df['clicks'] = df[['clicks', 'impressions']].min(axis=1)

    df['CTR'] = df['clicks'] / df['impressions'] * 100
    df['conversion'] = (df['purchases'] / df['impressions']) * 100

    return df

def generate_synthetic_data(n_control=14000, n_test=14000):
    np.random.seed(42)

    n = n_control + n_test
    user_ids = np.arange(1, n + 1)

    groups = ['control'] * n_control + ['test'] * n_test
    np.random.shuffle(groups)

    data = {
        'user_id': user_ids,
        'group': groups,
        'impressions': np.random.randint(6, 3000, size=n),
        'clicks': np.random.randint(0, 3000, size=n),
        'purchases': np.random.randint(0, 10, size=n),
        'session_length': np.random.randint(1, 90, size=n),
        'pages_per_session': np.random.randint(1, 5, size=n),
        'added_to_cart': np.random.randint(0, 15, size=n),
        'added_to_favorites': np.random.randint(0, 20, size=n),
        'bounces': np.random.randint(0, 2, size=n)
    }

    df = pd.DataFrame(data)

    df['purchases'] = df[['purchases', 'impressions']].min(axis=1)
    df['clicks'] = df[['clicks', 'impressions']].min(axis=1)

    df['CTR'] = df['clicks'] / df['impressions'] * 100
    df['conversion'] = (df['purchases'] / df['impressions']) * 100

    male_ratio = 0.7

    n_male_control = int(male_ratio * n_control)
    n_female_control = n_control - n_male_control
    n_male_test = int(male_ratio * n_test)
    n_female_test = n_test - n_male_test

    df_control = df[df['group'] == 'control'].copy()
    df_test = df[df['group'] == 'test'].copy()

    df_control['gender'] = ['male'] * n_male_control + ['female'] * n_female_control
    df_test['gender'] = ['male'] * n_male_test + ['female'] * n_female_test

    df = pd.concat([df_control, df_test])

    return df

df = generate_synthetic_data()
df_control, df_test = tune_control(df.copy()), tune_test(df.copy())
print(df_control.describe())
print(df_test.describe())
save_to_excel(df, df_control, df_test)
print("Файл успешно создан с тремя страницами.")