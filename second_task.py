import os, datetime, pandas as pd
from shutil import copyfile

def replace_files():
    """
    Function to replace files in 3 ways:
    1) D column in Excel file == 1
    2) Last two numbers in KA code are equal
    3) Date is between 20.06.2020 and 10.07.2020
    If some of conditions are True, replace in all ways
    """
    # Make a list of all files, created at prev program
    files = os.listdir("files")

    # Open Excel file
    xl = pd.read_excel('./Задание.xlsx')
    df = (xl[xl['Отправить'] == 1])
    # df.apply(move_file, axis=1, args=("Папка1",))
    
    # For all files in directory
    for i in files:
        if i.endswith(".txt"):
            # If is_userd, remove file in the end
            is_used = False

            # split filename by "_"
            num = i.split("_")

            # first way
            if df[df["Код_КА"] == "%s_%s" % (num[1],num[2])].values is not []:
                is_used = True
                move_file(i, 'Папка1')

            # second way
            if num[2][-1] == num[2][-2]:
                is_used = True
                move_file(i, 'Папка2') 

            # third way
            try:
                date = datetime.datetime.strptime(num[4][0:-4], '%d.%m.%Y')
            except ValueError as e:
                # print(e)
                date = datetime.datetime.strptime(
                    '%s.%s.%s' % (str(int(num[4][0:2]) - 1), num[4][3:5], num[4][-8:-4]),
                    '%d.%m.%Y')
            if  (
                date > datetime.datetime.strptime('20.06.2020', '%d.%m.%Y') and 
                date < datetime.datetime.strptime('10.07.2020', '%d.%m.%Y')
                ):
                is_used = True
                move_file(i, "Папка3")
            
            # Remove if any conditions were True
            if is_used:
                os.remove("./files/%s" % (i,))


def move_file(x, move_to):
    """
    Function to replace files
    x => filename or row from pd.DataFrame
    move_to => directory
    """
    # If there are no such directory, create one
    if not os.path.isdir(move_to):
        os.mkdir(move_to)

    # If x is row from pd.DataFrame
    if type(x) ==  pd.core.series.Series:
        for j in x[1].split(","):
            filename = "КА_%s_%s_%s.txt" % (x[0], j, x[2])
    else: # if x is a filename
        filename = x

    # Trying to replace file
    try:
        copyfile("./files/%s" %(filename,), "./%s/%s" % (move_to,filename))
    except FileNotFoundError as e:
        # print(e)
        pass


if __name__ == "__main__":
    replace_files()