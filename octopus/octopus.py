import xlwings as xw
from anytree import Node, RenderTree
from anytree.exporter import DotExporter
# from anytree.dotexport import RenderTreeGraph
import win32api         # for message box
import pandas as pd
from input import *


# ================================================MAIN================================================================
# @xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code


    sht_main = wb.sheets[sht_name_main]
    sht_run = wb.sheets[sht_name_run]

    # -------------------------------------------------------------------------------------
    df_main = sht_main.range('A1').options(
    pd.DataFrame, header=1, index=False, expand='table'
    ).value                                                         # fetch the data from sheet- 'Sheet1'

    root_cell_val = sht_run.range('F3').value     # from user entry --> int --> str
    index_no_list = df_main[df_main['Root'] == root_cell_val].index.tolist()     # returns a list with indices matching the search item in dataframe column

    # -------------------------------------------------------------------------------------
    if sht_run.range('F3').value is None:   # if a cell value is empty
        win32api.MessageBox(wb.app.hwnd, "Please enter the root no.", "Enter root no.")

    elif len(index_no_list) > 0:       # if the root entry is found
        root = Node(str(int(root_cell_val)))

        # root = Node("1")   
        n1 = Node(str(df_main.iloc[index_no_list[0], 1]), parent=root)
        n21 = Node(str(df_main.iloc[index_no_list[0], 2]), parent=n1)
        n22 = Node(str(df_main.iloc[index_no_list[0], 3]), parent=n1)
        n23 = Node(str(df_main.iloc[index_no_list[0], 4]), parent=n1)
        n24 = Node(str(df_main.iloc[index_no_list[0], 5]), parent=n1)
        
        DotExporter(root).to_dotfile("octopus.dot")    # create the dot file for image creation using Graphviz
        DotExporter(root).to_picture("octopus.png")

        # RenderTreeGraph(root).to_picture("octopus.png")    # create the png file using dot file
    
    else:               # if the root entry is NOT found
        win32api.MessageBox(wb.app.hwnd, "SORRY! the entered root no. is not found. \nPlease try again..", "Root no. NOT found")



if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()
