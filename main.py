"""
Generate subscript of image
"""

import numpy as np
import pandas as pd

# import docx2txt     # Just single string, not really interesting
from docx2python import docx2python


# Read the doc
def doc_reader(filepath):
    """
    I'm expecting 5 columns: TITEL	DATUM	TYPE EN FORMAAT	BEWAARPLAATS/EIGENAAR	AFBEELDING
    :return:
    """
    #

    if 0:
        # read in word file
        result = docx2txt.process(filepath)

    result = docx2python(filepath)

    body = result.body

    """
    len(body) = 2           # Probably different tables
    len(body[0]) = 162      # Rows
    len(body[0][0]) = 6     # Columns
    len(body[0][0][0]) = 1  # Rows in this column?
    len(body[0][0][0][0]) = 10 # Word/sentence?
    """

    return body     # lists the tables, rows, columns, enters, "letters" (but this is just a string)


def single_line(b, s=';'):
    if isinstance(b, list):
        return s.join([single_line(b_i).strip() for b_i in b])
    else:
        return b


# Generate table (pandas?)
def table_gen(body):

    d = []

    titles = [a[0].lower().strip() for a in body[0][0]]
    if titles[0] == '':   # index
        titles[0] = 'index'

    i_naam = titles.index('kunstenaar')
    i_titel = titles.index('titel')
    i_index = 0
    i_datum = titles.index('datum')
    i_typeformaat = titles.index('type en formaat')
    i_afbeelding = titles.index('afbeelding')     # Ignore!

    def check_if_title(l):
        return single_line(l[1]).lower() == titles[1]

    assert single_line(body[1]) == '', body[1]

    for i, l in enumerate(body[0]):

        if check_if_title(l):  #
            continue

        else:

            row = {}
            for i_c, c in enumerate(l):

                if 0 and (len(c) > 1):
                    print(c)  # Enters

                if i_c in (i_naam, i_titel):    # naam
                    row_i = c[0]    # Only before enter
                    # Remove "birth/death"
                    row_i = row_i.split("(")[0].strip()

                elif i_c == i_index:    # Index!
                    row_i = int(single_line(c))

                elif i_c in (i_datum, i_typeformaat):
                    row_i = single_line(c, ' ').lower()  # Shouldn't contain higher cases!

                elif i_c == i_afbeelding:
                    continue    # Do nothing

                else:
                    # Should be single line

                    row_i = single_line(c, ' ')

                row[titles[i_c]] = row_i

            d.append(row)

    df = pd.DataFrame(d, columns=titles)

    # Check if index is ok:
    for i, df_i in df.iterrows():
        if  df_i['index'] != i + df['index'][0]:
            print("at", df_i['index'], 'should be:', i + df['index'][0])

    # TODO check alphabetical
    list_lastfirstname = [a.lower() for a in list(df['kunstenaar'])]
    list_lastfirstname_sorted = sorted(list_lastfirstname)
    # To lower case

    for i, a in enumerate(list_lastfirstname):
        if a == '':
            print(i, a)

    for i, (a, b) in enumerate(zip(list_lastfirstname, list_lastfirstname_sorted)):
        if a != b:
            print("Error")
            print(i+1, a, '!=', b)

    # print(df)

    l_titles = []
    for i, df_i in df.iterrows():
        try:
            s = caption_gen(df_i) # df.iloc[0])
        except Exception as e:
            print(e)
            s = 'FAIL'

        assert ';' not in s, s

        l_titles.append({'index': i+1, 'onderschrift':s})

    df_titles = pd.DataFrame(l_titles)
    # df_titles.to_csv('C:/Users/Laurens_laptop_w/Downloads/corpus.csv', index=True, sep=';')
    df_titles.to_excel(r'C:/Users/Laurens_laptop_w/Downloads/corpus.xlsx', index=False)

    return df


def caption_gen(x, debug=False):

    if debug:
        print(x)

    # First name, last name
    s_name = x['kunstenaar']

    if s_name.strip() == '':
        print("Failed to make title", x)
        return ''  # Early stop

    s_name0 = s_name.split(';')[0]
    name_last, name_first = map(str.strip, s_name0.split('(')[0].split(',', 1))
    s0 = name_first + ' ' + name_last

    s_titel = x['titel']

    s_type = x['type en formaat']

    s_type.split(',')

    if s_type == '/':
        s_mat_formaat = 'afmetingen onbekend'
    else:
        s_mat_formaat = s_type

    s_bewaarplaats = x['bewaarplaats/eigenaar']
    if s_bewaarplaats == '/':
        s_col = 'bewaarplaats onbekend'
    else:
        s_col = s_bewaarplaats

    x_dat = x['datum'].strip()
    s_datum = 'datum onbekend' if x_dat == 's.d.' else x_dat

    # Naam, titel, datum, materiaal, formaat, collectie
    s = f'{s0}, {s_titel}, {s_datum}, {s_mat_formaat}, {s_col}'

    # print(s)

    assert ';' not in s, s

    return s


if __name__ == '__main__':
    f1 = 'CorpusKunstwerkenPerKunstenaarV2 (1).docx'
    f2 = 'CorpusOverzichtKunstwerkenV4_2.docx'
    # TODO also 'per_kunstenaar'
    body = doc_reader('C:/Users/Laurens_laptop_w/Downloads/' + f2)

    df = table_gen(body)
