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

            row = []
            for c in l:

                if 0 and (len(c) > 1):
                    print(c)  # Enters

                row_i = ';'.join(c)

                row.append(row_i)

            d.append(row)

        # TODO remove
        # if i >= 2:
        #     break

    df = pd.DataFrame(d, columns=titles)

    print(df)

    body[1] # TODO

    # for

    caption_gen(df.iloc[0])

    return df


def caption_gen(x):

    print(x)

    s = ''

    # First name, last name
    s_name = x['KUNSTENAAR']
    name_last, name_first = map(str.strip, s_name.split(';')[0].split('(')[0].split(',', 1))
    s0 = name_first + ' ' + name_last

    s += s0

    print(s)

    return s


if __name__ == '__main__':
    body = doc_reader('C:/Users/Laurens_laptop_w/Downloads/CorpusOverzichtKunstwerkenV2.docx')

    df = table_gen(body)
