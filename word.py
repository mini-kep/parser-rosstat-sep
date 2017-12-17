"""Reads data from tables in Word document. 
   Windows-only, requires MS Word installed.

   API:
      https://msdn.microsoft.com/en-us/library/office/ff837519.aspx
   Examples:
      http://stackoverflow.com/questions/10366596/reading-table-contetnts-in-ms-word-file-using-python

"""

import csv

ENCODING = 'utf8'

# -------------------------------------------------------------------------------
#
#     Application management
#
# -------------------------------------------------------------------------------


def win32_word_dispatch():
    # Lazy import of win32com - do not load Windows/MS Office libraries when
    # they are not called.
    import win32com.client as win32
    word = win32.Dispatch("Word.Application")
    word.Visible = 0
    return word


def open_ms_word():
    try:
        return win32_word_dispatch()
    except BaseException:
        raise Exception(
            "Apparently not a Windows machine or no MS Word installed.")


def close_ms_word(app):
    app.Quit()
    # ISSUE: must also quit somewhere by calling app.Quit()
    # like in
    # http://bytes.com/topic/python/answers/23946-closing-excel-application


def open_doc(path, word):
    """Opens doc file with win32com.client.

    Args:
        path: path to doc file.
        word: win32com.client.Dispatch object

    Yields:
        list of lists (rows)
    """
    from pywintypes import com_error
    try:
        word.Documents.Open(path)
    except com_error as e:
        if e.excepinfo[5] == -2146823683:
            print('\nEnsure word document document is not opened '
                  'in word already\n')
        elif e.excepinfo[5] == -2146823114:
            print('\nCheck path to doc file.\n'
                  'If the path you provided is relative, '
                  'pay attention to the fact that default directory '
                  'is system32\n')
        raise e
    return word.ActiveDocument


def get_table_count(doc):
    return doc.Tables.count


# -------------------------------------------------------------------------------
#
#     Cell value filter
#
# -------------------------------------------------------------------------------


def delete_double_space(line):
    return " ".join(line.split())


SPACE = " "
VOID = ""
APOCHAR = '"'
REPLACEMENTS = [('\r\x07', VOID)  # delete this symbol
                , ('\x0c', SPACE)  # sub with space
                , ('\x0b', SPACE)  # sub with space
                , ('\r', SPACE)  # sub with space
                , ("\u201c", APOCHAR), ("\u201d", APOCHAR), ('\x00', VOID)
                ]


def filter_cell_contents(cell_value):
    for a, b in REPLACEMENTS:
        cell_value = cell_value.replace(a, b)
    cell_value = delete_double_space(cell_value.strip())
    return cell_value


def get_filtered_cell_value(table, i, j):
    val = get_cell_value(table, i, j)
    return filter_cell_contents(val)


# -------------------------------------------------------------------------------
#
#     Word table iterators
#
# -------------------------------------------------------------------------------

def get_cell_value(table, i, j):
    try:
        return table.Cell(Row=i, Column=j).Range.Text
    # FIXME: which specific exceptions can it throw?
    except Exception:
        return ""


def cell_iter(table):
    for i in range(1, table.rows.count + 1):
        for j in range(1, table.columns.count + 1):
            yield i, j, get_filtered_cell_value(table, i, j)


def row_iter(table):
    for i in range(1, table.rows.count + 1):
        row = []
        for j in range(1, table.columns.count + 1):
            row = row + [get_filtered_cell_value(table, i, j)]
        yield row


# -------------------------------------------------------------------------------
#
#     Document-level iterators for .doc files
#
# -------------------------------------------------------------------------------

def query_all_tables(path, func):
    """Queries data from all tables within doc file.

    Args:
        path: path to doc file.
        func: name of the function used to parse the table

    Yields:
        list of lists (rows)
    """
    word = open_ms_word()
    doc = open_doc(path, word)
    total_tables = get_table_count(doc)
    for i, table in enumerate(doc.Tables):
        print("Reading table {} of {}...".format(i + 1, total_tables))
        yield func(table)
    close_ms_word(word)


def yield_continious_rows(path):
    """Yields rows of the table within doc file.

    Args:
        path: path to doc file.

    Yields:
        list of strings (row elements)
    """
    for y in query_all_tables(path, func=row_iter):
        for row in y:
            yield row
            
# -----------------------------------------------------------------------------
#
#    Write CSV
#
# -----------------------------------------------------------------------------


def to_csv(gen, csv_path):
    """Accept iterable of rows and write to *csv_path*."""
    with open(csv_path, 'w', encoding=ENCODING) as csvfile:
        filewriter = csv.writer(csvfile, delimiter='\t', lineterminator='\n')
        for row in gen:
            filewriter.writerow(row)
            
            
def from_csv(csv_path):
    """Yield iterable of rows from *csv_path*."""
    with open(csv_path, encoding=ENCODING) as csvfile:
        filereader = csv.reader(csvfile, delimiter='\t', lineterminator='\n')
        for row in filereader:
            yield(row)            

# -----------------------------------------------------------------------------
#
#    Interface
#
# -----------------------------------------------------------------------------
            
def doc2csv(doc_path,  csv_path):
    # consume Path instances
    doc_path, csv_path = str(doc_path), str(csv_path)
    # get list of rows from csv
    gen = yield_continious_rows(doc_path)
    # write generator to file 
    to_csv(gen, csv_path)
            
if __name__ == "__main__":
    pass