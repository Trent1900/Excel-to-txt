import os, openpyxl

def makeDir(filedir):
    for dir in os.listdir(filedir):
        if dir.endswith('.xlsm') or dir.endswith('.xlsx'):

            fullpath = os.path.join(filedir, dir)  # fullpath come with .xlxm
            filename = os.path.splitext(fullpath)[0]  # filename without .xlxm
            outDir = os.path.splitext(dir)[0]

            if not os.path.exists(outDir):
                os.makedirs(outDir)
            txtcon(fullpath,outDir)


def txtcon(fullpath,outDir):
    wb = openpyxl.load_workbook(fullpath)
    ws = wb['Display']
    maxrow = ws.max_row + 1
    maxcol = ws.max_column + 1

    list = []
    list2 = []
    #k is the number of Drawing in each FCS
    k=0

    for col in range(1, maxcol):

        for row in range(1, maxrow):
            a = ws.cell(row=row, column=col).value
            if a is not None:
                list.append(a)

        #consider only those collums with contents
        if len(list):
            # get the 1st element in the list as the name of the txt file
            i_str = str(list[0])
            # Create the new file name in the file path
            location = str(os.path.join(outDir, i_str))
            file_name = location + '.txt'

            # since we take the data from the 2nd element, 1st element is the tile of txt file
            # so we create a new list to store the data
            for i in range(0, len(list) - 1):
                if len(list) :
                    b = list[i + 1]
                    list2.append(b)

            # add a "return" ASCII to the each element of list2 for easy reading
            c = "\n".join(str(i) for i in list2)
            #'join'function will return a string, so c will be a string now.
            # c.strip()will consider those none empty string
            if c.strip():
                with open(file_name, 'w') as f:
                    f.write(str(c))
                #for each write to txt, k increase by 1
                k=k+1
            list2.clear()
            list.clear()

    print('for %s total %d file processed' % (outDir, k))

def main():
    filepath='target'
    makeDir(filepath)


if __name__ == "__main__":
    main()