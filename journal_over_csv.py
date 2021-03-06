import pandas as pd
import glob
import os
import warnings
import re

warnings.filterwarnings(action='ignore')

def reset_sr_index(sr):
    sr = sr.reset_index(drop=True)


print('converting journal_over excel to csv')
df_list = list()

df_start = str()
df_end = str()
df_cnt = 1

if not os.path.exists('csv/journal_over_csv'):
        os.mkdir('csv/journal_over_csv')

for f in glob.glob('./journal_over/*.xls'):

    df = pd.read_excel(f).T.reset_index().T.reset_index(drop=True)
    print(f, 'opened...')

    title = pd.Series() # 제목
    author = pd.Series() # 저자
    journal_name = pd.Series() # 학술지명
    doc_type = pd.Series() # 자료유형
    page = pd.Series() # 수록면
    publisher = pd.Series() # 발행처
    published_year = pd.Series() # 발행년도
    issues = pd.Series() # 권호사항
    issn = pd.Series() # ISSN
    subject = pd.Series() # 주제어

    check_page = False
    check_publisher = False
    check_issues = False
    check_published_year = False
    check_issn = False
    check_subject = False
    cnt = 0

    for i in range(len(df)):

        if df.iloc[i][0] == "제목" or df.iloc[i][0] == "서명":
            check_page = False
            check_publisher = False
            check_published_year = False
            check_issues = False
            check_issn = False 
            check_subject = False
            cnt += 1

            title.set_value(cnt, df.iloc[i][1])
            
        if df.iloc[i][0] == "저자":
            author.set_value(cnt, df.iloc[i][1])
            
        if df.iloc[i][0] == "학술지명":
            journal_name.set_value(cnt, df.iloc[i][1])
        
        if df.iloc[i][0] == "자료유형":
            doc_type.set_value(cnt, df.iloc[i][1])
            
        if df.iloc[i][0] == "수록면":
            page.set_value(cnt, df.iloc[i][1])
            check_page = True
        if not check_page:
            page.set_value(cnt, None)
            
        if df.iloc[i][0] == "발행처":
            publisher.set_value(cnt, df.iloc[i][1])
            check_publisher = True
        if not check_publisher:
            publisher.set_value(cnt, None)
            
        if df.iloc[i][0] == "발행년도" or df.iloc[i][0] == "출판년":
            published_year.set_value(cnt, df.iloc[i][1])
            check_published_year = True
        if not check_published_year:
            published_year.set_value(cnt, None)
            
        if df.iloc[i][0] == "권호사항":
            issues.set_value(cnt, df.iloc[i][1])
            check_issues = True
        if not check_issues:
            issues.set_value(cnt, None)
            
        if df.iloc[i][0] == "ISSN":
            issn.set_value(cnt, df.iloc[i][1])
            check_issn = True
        if not check_issn:
            issn.set_value(cnt, None)
            
        if df.iloc[i][0] == "주제어":
            subject.set_value(cnt, df.iloc[i][1])
            check_subject = True
        if not check_subject:
            subject.set_value(cnt, None)

    reset_sr_index(title)
    reset_sr_index(author)
    reset_sr_index(journal_name)
    reset_sr_index(doc_type)
    reset_sr_index(page)
    reset_sr_index(publisher)
    reset_sr_index(published_year)
    reset_sr_index(issues)
    reset_sr_index(issn)
    reset_sr_index(subject)


    new_df = pd.DataFrame(
                {
                    '제목': title,
                    '저자': author,
                    '학술지명': journal_name,
                    '자료유형': doc_type,
                    '수록면': page,
                    '발행처': publisher,
                    '발행년도': published_year,
                    '권호사항': issues,
                    'ISSN': issn,
                    '주제어': subject
                }
        ,
        columns = ['제목', '저자', '학술지명', '자료유형', '수록면', '발행처', '발행년도', '권호사항', 'ISSN', '주제어']
                )
    print('df of ' + f + ' complete')
    df_list.append(new_df)
    print(str(df_cnt%100) + '/100 apppending to df_list')

    if df_cnt % 100 == 1:
        df_start = re.findall("\d+", str(f))[0]
    if (df_cnt % 100 == 0) or (f == glob.glob('./journal_over/*.xls')[-1]):
        df_end = re.findall("\d+", str(f))[0]
        pd.concat(df_list).to_csv('csv/journal_over_csv/journal_over_' + df_start + '-' + df_end + '.csv', sep=',')
        print('journal_over_' + df_start + '-' + df_end + '.csv saved')
        df_list = []
    df_cnt += 1
    
print('journal_over complete')