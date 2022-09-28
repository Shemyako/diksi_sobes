import pandas as pd, os

def make_files():
    """
    Function to open Excel file and 
    run create_file function for all of them
    """
    # Creating dir for new files
    if os.path.isdir("files"):
        os.rmdir("files")
    os.mkdir("files")

    # Open Excel file
    xl = pd.read_excel('./Задание.xlsx')
    # Run create_function for every rows
    xl.apply(create_file, axis=1)


def create_file(x):
    """
    Function to create file
    """
    # For all documents in B column
    for i in x[1].split(","):
        # Create file КА_<Код КА>_<тип_документа>_<дата>.txt
        with open("./files/КА_%s_%s_%s.txt" % (x[0], i, x[2]), "w"):
            ### Write something into file
            pass


if __name__ == "__main__":
    make_files()