import pandas as pd
import glob
import warnings

warnings.filterwarnings(action='ignore')

def reset_sr_index(sr):
    sr = sr.reset_index(drop=True)

def convert_to_csv():
    print('converting degree excel to csv')
    df_list = list()

    for f in glob.glob('./degree/*.xls'):

        df = pd.read_excel(f)
        print(f, 'opened...')

        title = pd.Series() # 서명
        author = pd.Series() # 저자
        abstract = pd.Series() # 초록
        chapter = pd.Series() # 목차
        degree = pd.Series() # 학위논문사항
        publisher = pd.Series() # 발행사항
        year = pd.Series() # 출판년
        KDC = pd.Series() # KDC
        subject = pd.Series() # 주제어

        check_abs = False
        check_abs_a = False
        check_chapter = False
        check_degree = False
        check_publisher = False
        check_year = False
        check_KDC = False
        check_subj = False
        cnt = 0

        for i in range(len(df)):
    
            if df.iloc[i][0] == "제목" or df.iloc[i][0] == "서명":
                check_abs = False
                check_abs_a = False
                check_chapter = False
                check_degree = False
                check_publisher = False
                check_year = False
                check_KDC = False
                check_subj = False
                
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
                
            if df.iloc[i][0] == "목차":
                chapter.set_value(cnt, df.iloc[i][1])
                check_chapter = True
            if not check_chapter:
                chapter.set_value(cnt, None)
                
            if df.iloc[i][0] == "학위논문사항":
                degree.set_value(cnt, df.iloc[i][1])
                check_degree = True
            if not check_degree:
                degree.set_value(cnt, None)
                
            if df.iloc[i][0] == "발행사항":
                publisher.set_value(cnt, df.iloc[i][1])
                check_publisher = True
            if not check_publisher:
                publisher.set_value(cnt, None)
                
            if df.iloc[i][0] == "출판년":
                year.set_value(cnt, df.iloc[i][1])
                check_year = True
            if not check_year:
                year.set_value(cnt, None)
            if df.iloc[i][0] == "KDC":
                KDC.set_value(cnt, df.iloc[i][1])
                check_KDC = True
                
            if not check_KDC:
                KDC.set_value(cnt, None)
                
            if df.iloc[i][0] == "주제어":
                subject.set_value(cnt, df.iloc[i][1])
                check_subj = True
            if not check_subj:
                subject.set_value(cnt, None)

        reset_sr_index(title)
        reset_sr_index(author)
        reset_sr_index(abstract)
        reset_sr_index(degree)
        reset_sr_index(publisher)
        reset_sr_index(year)
        reset_sr_index(KDC)
        reset_sr_index(abstract)


        new_df = pd.DataFrame(
                    {
                        '제목': title,
                        '저자': author,
                        '초록': abstract,
                        '목차': chapter,
                        '학위논문사항': degree,
                        '발행사항': publisher,
                        '출판년': year,
                        'KDC': KDC,
                        '주제어': subject
                    }
            ,
            columns = ['제목', '저자', '초록', '목차', '학위논문사항', '발행사항', '출판년', 'KDC', '주제어']
                    )
        print('df of ' + f + ' complete')
        df_list.append(new_df)
        print('apppending to full list')
    print('done')
    full_df = pd.concat(df_list)
    return full_df.reset_index(drop=True)