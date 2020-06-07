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


# Generate table (pandas?)
def table_gen(body):

    d = []

    for i, l in enumerate(body[0]):
        if i == 0:
            titles = [c[0] for c in l]

            # TODO check when it is title

            continue

        else:

            row = {}
            for i_c, c in enumerate(l):

                if 0 and (len(c) > 1):
                    print(c)  # Enters

                if i_c == 0:    # naam
                    row_i = c[0]    # Only before enter
                else:
                    row_i = ';'.join(c)

                row[titles[i_c]] = row_i

            d.append(row)

        # TODO remove
        # if i >= 2:
        #     break

    df = pd.DataFrame(d, columns=titles)

    print(df)

    body[1] # TODO

    # for

    l_titles = []
    for i, df_i in df.iterrows():
        try:
            s = caption_gen(df_i) # df.iloc[0])
        except Exception as e:
            print(e)
            s = 'FAIL'
        l_titles.append(s)
    df_titles = pd.DataFrame(l_titles)
    df_titles.to_csv('C:/Users/Laurens_laptop_w/Downloads/corpus.csv', index=True, sep=';')

    return df


def caption_gen(x, debug=False):

    if debug:
        print(x)

    # First name, last name
    s_name = x['KUNSTENAAR']

    if s_name.strip() == '':
        print("Failed to make title", x)
        return ''  # Early stop

    s_name0 = s_name.split(';')[0]
    name_last, name_first = map(str.strip, s_name0.split('(')[0].split(',', 1))
    s0 = name_first + ' ' + name_last

    s_titel = x['TITEL']

    s_type = x['TYPE EN FORMAAT']

    s_type.split(',')

    if s_type == '/':
        s_mat_formaat = 'afmetingen onbekend'
    else:
        s_mat_formaat = s_type

    s_bewaarplaats = x['BEWAARPLAATS/EIGENAAR']
    if s_bewaarplaats == '/':
        s_col = 'bewaarplaats onbekend'
    else:
        s_col = s_bewaarplaats

    x_dat = x['DATUM'].strip()
    s_datum = 'datum onbekend' if x_dat == 's.d.' else x_dat

    # Naam, titel, datum, materiaal, formaat, collectie
    s = f'{s0}, {s_titel}, {s_datum}, {s_mat_formaat}, {s_col}'

    print(s)

    return s


if __name__ == '__main__':
    body = doc_reader('C:/Users/Laurens_laptop_w/Downloads/CorpusOverzichtKunstwerkenV2.docx')

    df = table_gen(body)
