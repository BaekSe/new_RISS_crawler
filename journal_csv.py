import pandas as pd
import glob
import warnings

warnings.filterwarnings(action='ignore')

def convert_to_csv():
    print('converting journal_kr excel to csv')
    df_list = list()

    for f in glob.glob('./journal_kr/*.xls'):

        df = pd.read_excel(f).T.reset_index().T.reset_index(drop=True)
        print(f, 'opened...')

        title = pd.Series() # 서명
        author = pd.Series() # 저자
        published_year = pd.Series() # 출판년
        publish_info = pd.Series() # 발행사항
        chapter = pd.Series() # 목차
        doc_type = pd.Series() # 자료유형
        kdc = pd.Series() # KDC
        isbn = pd.Series() # ISBN
        subject = pd.Series() # 주제어

        check_chapter = False
        check_KDC = False
        check_ISBN = False
        check_subj = False
        cnt = 0

        for i in range(len(df)):
    
            if df.iloc[i][0] == "제목" or df.iloc[i][0] == "서명":
                check_chapter = False
                check_KDC = False
                check_ISBN = False
                check_subj = False
                cnt += 1

                title.set_value(cnt, df.iloc[i][1])
                
            if df.iloc[i][0] == "저자":
                author.set_value(cnt, df.iloc[i][1])
                
            if df.iloc[i][0] == "출판년" or df.iloc[i][0] == "발행년도":
                published_year.set_value(cnt, df.iloc[i][1])
                
            if df.iloc[i][0] == "발행사항":
                publish_info.set_value(cnt, df.iloc[i][1])
                
            if df.iloc[i][0] == "목차":
                chapter.set_value(cnt, df.iloc[i][1])
                check_chapter = True
            if not check_chapter:
                chapter.set_value(cnt, None)
                
            if df.iloc[i][0] == "자료유형":
                doc_type.set_value(cnt, df.iloc[i][1])
                
            if df.iloc[i][0] == "KDC":
                kdc.set_value(cnt, df.iloc[i][1])
                check_KDC = True
            if not check_KDC:
                kdc.set_value(cnt, None)
                
            if df.iloc[i][0] == "ISBN":
                isbn.set_value(cnt, df.iloc[i][1])
                check_ISBN = True
            if not check_ISBN:
                isbn.set_value(cnt, None)
                
            if df.iloc[i][0] == "주제어":
                subject.set_value(cnt, df.iloc[i][1])
                check_subj = True
            if not check_subj:
                subject.set_value(cnt, None)

        title = title.reset_index(drop=True)
        author = author.reset_index(drop=True)
        published_year = published_year.reset_index(drop=True)
        publish_info = publish_info.reset_index(drop=True)
        chapter = chapter.reset_index(drop=True)
        doc_type = doc_type.reset_index(drop=True)
        kdc = kdc.reset_index(drop=True)
        isbn = isbn.reset_index(drop=True)
        subject = subject.reset_index(drop=True)


        new_df = pd.DataFrame(
                    {
                        '서명': title,
                        '저자': author,
                        '발행년도': published_year,
                        '발행정보': publish_info,
                        '목차': chapter,
                        '자료유형': doc_type,
                        'KDC': kdc,
                        'ISBN': isbn,
                        '주제어': subject
                    }
            ,
            columns = ['서명', '저자', '발행년도', '발행정보', '목차', \
            '자료유형', 'KDC', 'ISBN', '주제어']
                    )
        print('df of ' + f + ' complete')
        df_list.append(new_df)
        print('apppending to full list')
    print('done')
    full_df = pd.concat(df_list)
    return full_df.reset_index(drop=True)
    #full_df.to_csv('book.csv', sep=',', na_rep='NaN', encoding = 'utf-8')
