import os
import xlsxwriter


def fasta_reader_hieu(filename):
    '''
    This is hieu's own fasta reader. The input should include just a reader and a sequence follow up
    example:
        >Name
        Sequence
    '''
    seq_list = []

    with open(filename+".txt", "r") as file:
        count = 0
        for line in file:
            if line[0] == '>':
                seq_list.append([line[:-1][1:]])
            else:
                seq_list[count].append([line[:-1]])
                count+=1
    print('extracted :', seq_list)
    return seq_list


def delSpace(string):
    #for normalizing the sequence
    n = len(string)
    newStr = ""
    aa = ['D','T','S','E','P','G','A','C','V','M','I','L','Y','F','H','K','R','W','Q','N']
    for i in range(n):
        if string[i] in aa:
            newStr += string[i]
    return newStr


def hieu(filename, sequence, w=20):

    sequence = delSpace(str(sequence))
    filename = str(filename)+'.xlsx' if '.xlsx' or '.xls' not in filename else 0
    f = {'D':-1,'T':0,'S':0,'E':-1,'P':0,'G':0,'A':0,'C':-0.1,'V':0,'M':0,'I':0,'L':0,'Y':0,'F':0,'H':0.1,'K':1,'R':1,'W':0,'Q':0,'N':0}

    with xlsxwriter.Workbook(filename) as workbook:

        worksheet = workbook.add_worksheet('Data')
        worksheet.write(0,0,'index')
        worksheet.write(0,1,'charge')

        seq_limit = len(sequence)-(w-1)

        #writing the window index and the charge
        for i in range(seq_limit):
            sum = 0
            for j in range(i,i+w):
                sum += f[sequence[j]]
            worksheet.write(i+1,0,str(i+1) + "~" + str(i+w))
            worksheet.write(i+1,1,sum)

        #add smooth scatter chart
        chart = workbook.add_chart({'type': 'scatter',
                                    'subtype': 'smooth'})

        #add a single series with categories being the window and values being the charge
        chart.add_series({'name':       '=Data!$B$1',
                          'categories': '=Data!$A$2:$A${}'.format(seq_limit),
                          'values':     '=Data!$B$2:$B${}'.format(seq_limit),
                        })

        chart.set_x_axis({'name': 'Window index'})
        chart.set_y_axis({'name': 'Partial charge'})

        # Set an Excel chart style.
        chart.set_style(15)

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('D2', chart, {'x_offset': 25, 'y_offset': 10})

def fabia(seq_list, w=20):

    f = {'D':-1,'T':0,'S':0,'E':-1,'P':0,'G':0,'A':0,'C':-0.1,'V':0,'M':0,'I':0,'L':0,'Y':0,'F':0,'H':0.1,'K':1,'R':1,'W':0,'Q':0,'N':0}

    with xlsxwriter.Workbook("Summary.xlsx") as workbook:

        for i in range(len(seq_list)):

            seq_name = seq_list[i][0]
            sequence = str(seq_list[i][1])[1:][:-1]

            print(ADDED+sequence)

            worksheet = workbook.add_worksheet(seq_name)
            worksheet.write(0,0,'index')
            worksheet.write(0,1,'charge')

            seq_limit = len(sequence)-(w-1)

            #writing the window index and the charge
            for i in range(seq_limit):
                sum = 0
                for j in range(i,i+w):
                    sum += f[sequence[j]]
                worksheet.write(i+1,0,str(i+1) + "~" + str(i+w))
                worksheet.write(i+1,1,sum)

            #add smooth scatter chart
            chart = workbook.add_chart({'type': 'scatter',
                                        'subtype': 'smooth'})

            #add a single series with categories being the window and values being the charge
            chart.add_series({'name':       '='+seq_name+'!$B$1',
                              'categories': '='+seq_name+'!$A$2:$A${}'.format(seq_limit),
                              'values':     '='+seq_name+'!$B$2:$B${}'.format(seq_limit),
                            })

            chart.set_x_axis({'name': 'Window index'})
            chart.set_y_axis({'name': 'Partial charge'})

            # Set an Excel chart style.
            chart.set_style(15)

            # Insert the chart into the worksheet (with an offset).
            worksheet.insert_chart('D2', chart, {'x_offset': 25, 'y_offset': 10})


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' + directory)


def main():

    createFolder('analysis result')
    seq_list = fasta_reader_hieu('input_test')
    os.chdir("./analysis result")

    for i in range(len(seq_list)):
        hieu(seq_list[i][0], seq_list[i][1])

    fabia(seq_list)


if __name__=='__main__':
    main()
