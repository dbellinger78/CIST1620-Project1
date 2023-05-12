from gui import *


def main():
    """
    This project is based on Lab 5 where data was read from a PDF file and written to an Excel file. 
    This program reads data from an Excel spreadsheet and writes data to the GUI interface.
    """
    window = Tk()
    window.title('Excel File Reader')
    window.geometry('1600x1000')
    widgets = GUI(window)

 
    window.mainloop()


if __name__ == '__main__':
    main()
