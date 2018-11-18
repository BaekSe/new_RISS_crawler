import pandas as pd
import glob
import warnings

warnings.filterwarnings(action='ignore')

def convert_to_csv():
    print('converting book excel to csv')
    df_list = list()

    for f in glob.glob('./book/*.xls'):

        df = pd.read_excel(f)
        print(f, 'opened...')

        title = pd.Series() # 제목
        author = pd.Series() # 저자
        abstract = pd.Series() # 초록
        doc_type = pd.Series() # 자료유형
        book_name = pd.Series() # 학술지명
        page = pd.Series() # 수록면
        published_year = pd.Series() # 발행년도
        publisher = pd.Series() # 발행처
        issues = pd.Series() # 권호사항
        kdc = pd.Series() # KDC
        subject = pd.Series() # 주제어

        check_abs = False
        check_KDC = False
        check_subj = False
        check_page = False
        check_abs_a = False
        check_journal_name = False
        check_year = False
        check_publisher = False
        check_issues = False
        cnt = 0

        for i in range(len(df)):
    
            if df.iloc[i][0] == "제목" or df.iloc[i][0] == "서명":
                check_abs = False
                check_KDC = False
                check_subj = False
                check_page = False
                check_abs_a = False
                check_journal_name = False
                check_year = False
                check_publisher = False
                check_issues = False
                
                title.set_value(cnt, df.iloc[i][1])
                cnt += 1
                
            if df.iloc[i][0] == "저자":
                author.set_value(cnt, df.iloc[i][1])
                
            if df.iloc[i][0] == "초록" and (not check_abs):
                abstract.set_value(cnt, df.iloc[i][1])
                check_abs = True
                check_abs_a = True
            if not check_abs_a:
                abstract.set_value(cnt, None)
                
            if df.iloc[i][0] == "자료유형":
                doc_type.set_value(cnt, df.iloc[i][1])
                
            if df.iloc[i][0] == "학술지명":
                book_name.set_value(cnt, df.iloc[i][1])
                check_journal_name = True
            if not check_journal_name:
                book_name.set_value(cnt, None)
                
            if df.iloc[i][0] == "수록면":
                page.set_value(cnt, df.iloc[i][1])
                check_page = True
            if not check_page:
                page.set_value(cnt, None)
                
            if df.iloc[i][0] == "발행년도":
                published_year.set_value(cnt, df.iloc[i][1])
                check_year = True
            if not check_year:
                published_year.set_value(cnt, None)
                
            if df.iloc[i][0] == "발행처":
                publisher.set_value(cnt, df.iloc[i][1])
                check_publisher = True
            if not check_publisher:
                publisher.set_value(cnt, None)
                
            if df.iloc[i][0] == "권호사항":
                issues.set_value(cnt, df.iloc[i][1])
                check_issues = True
            if not check_issues:
                issues.set_value(cnt, None)
                
            if df.iloc[i][0] == "KDC":
                kdc.set_value(cnt, df.iloc[i][1])
                check_KDC = True
                
            if not check_KDC:
                kdc.set_value(cnt, None)
                
            if df.iloc[i][0] == "주제어":
                subject.set_value(cnt, df.iloc[i][1])
                check_subj = True
            if not check_subj:
                subject.set_value(cnt, None)

        title = title.reset_index(drop=True)
        author = author.reset_index(drop=True)
        abstract = abstract.reset_index(drop=True)
        doc_type = doc_type.reset_index(drop=True)
        book_name = book_name.reset_index(drop=True)
        page = page.reset_index(drop=True)
        published_year = published_year.reset_index(drop=True)
        publisher = publisher.reset_index(drop=True)
        issues = issues.reset_index(drop=True)
        kdc = kdc.reset_index(drop=True)
        subject = subject.reset_index(drop=True)


        new_df = pd.DataFrame(
                    {'제목': title,
                    '저자': author,
                    '초록': abstract, 
                    '자료유형': doc_type, 
                    '학술지명': book_name, 
                    '수록면': page,
                    '발행년도': published_year,
                    '발행처': publisher,
                    '권호사항': issues, 
                    'KDC': kdc,
                    '주제어': subject}
            ,
            columns = ['제목', '저자', '초록', '자료유형', '학술지명',\
                    '수록면', '발행년도', '발행처', '권호사항', 'KDC', '주제어']
                    )
        print('df of ' + f + ' complete')
        df_list.append(new_df)
        print('apppending to full list')
    print('done')
    full_df = pd.concat(df_list)
    return full_df.reset_index(drop=True)
    #full_df.to_csv('book.csv', sep=',', na_rep='NaN', encoding = 'utf-8')