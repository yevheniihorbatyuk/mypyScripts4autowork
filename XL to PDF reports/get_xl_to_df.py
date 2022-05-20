import pandas as pd

path = '../data/tpl_access_data.xlsx'
sheet_name = 'list'

df_excel = pd.read_excel(path, sheet_name='list')
df_shape = df_excel.shape


print('TABLE readied. Shape', df_shape)

# Data engineering
df_excel.sort_values(by=df_excel.columns[0], inplace=True)
df_excel.reset_index(inplace=True)
df_excel.drop(columns='index', inplace=True)


# LOG
print('Columns:', df_excel.columns)

df_excel.rename(columns={'ПІБ': 'Person',
                         'Прізвище': 'Surname',
                         'Імя': 'Name',
                         'По-батькові': 'FaName',
                         'Званння': 'Rank',
                         'Посада': 'Position',
                         'ІНП': 'VATIN',

                         'Номер': 'Phone',
                         'ВПС': 'Location',
                         'Статус': 'Status',
                         'ДІЯ': 'Action',
                         'Дата': 'Date',
                         'Номер наказу': 'DocNum',
                         'ЕЦП': 'EDK',
                         'БД': 'BD',

                         'Логін': 'Login',
                         'Пароль': 'Password',
                         'IP': 'IP'},
                inplace=True)

# LOG
print('\nColumns:', df_excel.columns)

v_count = df_excel['Person'].value_counts()

