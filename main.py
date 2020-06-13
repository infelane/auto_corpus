"""
Generate subscript of image
"""

import os

import numpy as np
import pandas as pd

# import docx2txt     # Just single string, not really interesting
from docx2python import docx2python
import openpyxl


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
        return s.join([single_line(b_i, s).strip() for b_i in b])
    else:
        return b


# Generate table (pandas?)
def table_gen(body):

    if len(body[0]) == 1:    # Painter grouped
        titles = [a[0].lower().strip() for a in body[1][0]]
    else:
        titles = [a[0].lower().strip() for a in body[0][0]]

    if titles[0] == '':   # index
        titles[0] = 'index'

    if 'kunstenaar' in titles:
        i_naam = titles.index('kunstenaar')
    else:
        i_naam = None

    i_titel = titles.index('titel')
    i_index = 0
    i_datum = titles.index('datum')
    i_typeformaat = titles.index('type en formaat')
    i_afbeelding = titles.index('afbeelding')     # Ignore!

    def check_if_title(l):
        return single_line(l[1]).lower() == titles[1]

    def body_work(body_i, author=None):
        d = []
        for i, l1 in enumerate(body_i):

            if check_if_title(l1):  #
                continue

            else:

                row = {}

                if author is not None:  # For the other file
                    row["kunstenaar"] = author

                for i_c, c in enumerate(l1):

                    if 0 and (len(c) > 1):
                        print(c)  # Enters

                    if i_c in (i_naam, i_titel):  # naam
                        row_i = c[0]  # Only before enter
                        # Remove "birth/death"
                        row_i = row_i.split("(")[0].strip()

                        if ((i_c == i_naam) and (author is None)) or \
                                ((i_c == i_titel) and (author is not None)):  # add everything after the name
                            # All the rest, to single

                            l1 = single_line(c[1:], ' ').strip()

                            # Handle html < > stuff!
                            l2 = ''
                            for l_i in l1.split("<"):
                                l2 += l_i.split(">", 1)[-1]

                            l1 = l2

                            row['verantwoording'] = remove_double_space(l1)

                    elif i_c == i_index:  # Index!
                        row_i = int(single_line(c))

                    elif i_c in (i_datum, i_typeformaat):
                        row_i = single_line(c, ' ').lower()  # Shouldn't contain higher cases!

                    elif i_c == i_afbeelding:
                        continue  # Do nothing

                    else:
                        # Should be single line

                        row_i = single_line(c, ' ')

                    row[titles[i_c]] = row_i

                d.append(row)

        return d

    d = []

    author = None
    for body_i in body:

        if len(body_i) == 1:    # Artists grouped:

            naam = single_line(body_i, '').lower()

            if naam == "":
                continue     # Empty

            if "henri" in naam:
                naam = "De Braekeleer, Henri"

            # TODO other 2
            elif "xavier" in naam:
                naam = "Mellery, Xavier"

            elif "smet" in naam:
                naam = "De Smet, Léon"

            else:
                raise ValueError(naam)

            author = naam
            continue

        d_i = body_work(body_i, author=author)
        d.extend(d_i)

    df = pd.DataFrame(d,
                      # columns=titles  # Might forget some extra keys if uncommenting
                      )

    return df


def check_alphabetical(df):
    # check alphabetical

    # To lower case
    list_lastfirstname = [a.lower() for a in list(df['kunstenaar'])]
    list_lastfirstname_sorted = sorted(list_lastfirstname)

    # Check empty fields
    for i, a in enumerate(list_lastfirstname):
        if a == '':
            print(i, a)

    for i, (a, b) in enumerate(zip(list_lastfirstname, list_lastfirstname_sorted)):
        if a != b:
            print("Error")
            print(i+1, a, '!=', b)


def processing(df):

    # Check if index is ok:
    for i, df_i in df.iterrows():
        if df_i['index'] != i + df['index'][0]:
            print("at", df_i['index'], 'should be:', i + df['index'][0])

    if 1:
        count_author(df)

    if 1:
        gen_verantwoording(df)

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
    df_titles.to_excel(os.path.join(folder, 'corpus.xlsx'), index=False)

    return df


def sort_df(df, column, key):
    '''Takes dataframe, column index and custom function for sorting,
    returns dataframe sorted by this column using this function'''

    temp = {i: key(df[column][i]) for i in df.index}

    return df.iloc[sorted(temp, key=lambda x: temp[x])]


def count_author(df):

    df_kunstenaar = sort_df(df, "kunstenaar", str.lower).groupby('kunstenaar', sort=False).size()

    df_kunstenaar.to_excel(os.path.join(folder, 'kunstenaars_tellen_base.xlsx'))


def gen_verantwoording(df):

    a = df['verantwoording']
    print(a)

    l_verantwoording_safe = []

    for _, a in df.iterrows():
        # print(a)

        s = f"catalogus nummer {a['index']}: {a['verantwoording']}"

        l_verantwoording_safe.append(s)

    df_verantwoording_safe = pd.DataFrame(l_verantwoording_safe)

    df_verantwoording_safe.to_excel(os.path.join(folder, 'cat_verantwoording.xlsx'), index=False, header=False)

    return

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


def remove_double_space(s):
    return ' '.join(s.split())


if __name__ == '__main__':
    folder = 'C:/Users/admin/Downloads' # 'C:/Users/Laurens_laptop_w/Downloads/'

    f1 = 'CorpusKunstwerkenPerKunstenaarV4.docx'   # 'CorpusKunstwerkenPerKunstenaarV2 (1).docx'
    f2 = 'CorpusOverzichtKunstwerkenV6.docx'   # 'CorpusOverzichtKunstwerkenV4_2.docx'
    # TODO also 'per_kunstenaar'

    body1 = doc_reader(os.path.join(folder, f1))
    df1 = table_gen(body1)

    # check_alphabetical(df1)

    body2 = doc_reader(os.path.join(folder, f2))
    df2 = table_gen(body2)

    check_alphabetical(df2)

    df = pd.concat([df2, df1], ignore_index=True)

    processing(df)

    # TODO generate df for both files
    # TODO join both files
