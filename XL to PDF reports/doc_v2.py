from docxtpl import DocxTemplate
from get_xl_to_df import df_excel, df_shape, v_count
from docx2pdf import convert
from tqdm import tqdm
from os import mkdir, rmdir
doc = DocxTemplate('../data/template_access_report.docx')

try:
    rmdir('../reports')
    rmdir('../reports/pdf')
    mkdir('../reports')
    mkdir('../reports/pdf')
except:
    print('DIR exist')

for person in tqdm(v_count.index, desc='Loading...'):
    # print('Person:', person)

    index = df_excel.index[df_excel['Person'] == person].to_list()
    # print('INDEX :', index)

    dbAccessTable = []
    condition = True
    for i in index:
        # print(i)
        dbAccessTable.append({"action": df_excel.loc[i, 'Action'],
                              "date": df_excel.loc[i, 'Date'],
                              "docNum": df_excel.loc[i, 'DocNum'],
                              "edk": df_excel.loc[i, 'EDK'],
                              "numPos ": df_excel.loc[i, 'Action'],
                              'db': df_excel.loc[i, 'BD'],
                              'logPas': str(df_excel.loc[i, 'Login']) + '; ' + str(df_excel.loc[i, 'Password'])
                              })

    context = {
        'name': person,
        'VATIN': df_excel.at[index[0], 'VATIN'],
        'position': df_excel.at[index[0], 'Position'],
        'place': df_excel.at[index[0], 'Location'],
        'ip': df_excel.at[index[0], 'IP'],

        "dbAccessTable": dbAccessTable,
    }

    # render context into the document object
    doc.render(context)

    # save the document object as a word file
    path = f'../reports/{person}.docx'
    doc.save(path)

    convert(path, path.replace(".docx", ".pdf").replace("reports/", "reports/pdf/"))
