import os
from threading import Thread
from statistics import Statistics


def find_excel():
    files, xslx_files = os.listdir(), []
    for file in files:
        if os.path.isfile(file):
            if os.path.splitext(file)[1] == '.xlsx' and not file.startswith(r'~$'):
                xslx_files.append(file)
    return xslx_files


def main():
    current_path = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_path)
    xlsx_files = find_excel()

    if len(xlsx_files) == 0:
        print('XLSX FILE NOT FOUND')
        exit()
    elif len(xlsx_files) == 1:
        statistics = Statistics()
        statistics.count(xlsx_files[0])
    else:
        # length = len(xlsx_files)
        # threads = [None] * length
        # for i in range(0, length):
        #     statistics = Statistics()
        #     threads[i] = Thread(target=statistics.count(),
        #                         args=[xlsx_files[i]])
        #     threads[i].start()
        for xlsx_file in xlsx_files:
            statistics = Statistics()
            os.chdir(current_path)
            statistics.count(xlsx_file)


if __name__ == '__main__':
    main()
