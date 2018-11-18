import journal_csv
import journal_over_csv
import book_csv
import degree_csv
import etc_csv
import pandas as pd

try:
    df = journal_csv.convert_to_csv()
    df.to_csv('journal.csv', sep=',')
except:
    pass


try:
    df = journal_over_csv.convert_to_csv()
    df.to_csv('journal_over.csv', sep=',')
except:
    pass


try:
    df = degree_csv.convert_to_csv()
    df.to_csv('degree.csv', sep=',')
except:
    pass


try:
    df = book_csv.convert_to_csv()
    df.to_csv('book.csv', sep=',')
except:
    pass


try:
    df = etc_csv.convert_to_csv()
    df.to_csv('etc.csv', sep=',')
except:
    pass
