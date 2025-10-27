import numpy as np
import pandas as pd
from openpyxl import load_workbook

def save_to_excel(df, filename="ficha_ab.xlsx"):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="ficha")

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

    # Увеличиваем длину сессии и уменьшаем количество покупок для тестовой группы
    df.loc[df['group'] == 'test', 'session_length'] = (df.loc[df['group'] == 'test', 'session_length'] * 1.5).astype(int)
    df.loc[df['group'] == 'test', 'purchases'] = (df.loc[df['group'] == 'test', 'purchases'] * 0.2).astype(int)

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
df.describe()
save_to_excel(df)
print("Файл успешно создан с одной страницей.")