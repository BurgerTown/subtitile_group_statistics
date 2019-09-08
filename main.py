import os
from threading import Thread
from statistics import Statistics


def find_excel():
    things, dirs, xslx_files = os.listdir(), [], []
    for sth in things:
        if os.path.isfile(sth):
            if os.path.splitext(sth)[1] == '.xlsx' and not sth.startswith(r'~$'):
                xslx_files.append(sth)
                print('ADDED TO THE LIST: {} '.format(sth))
        else:
            dirs.append(sth)
    return dirs, xslx_files


def exclude_done(dirs, xlsx_files):
    results = []
    for xlsx_file in xlsx_files:
        if xlsx_file == 'Template.xlsx':
            continue
        else:
            file_name = os.path.splitext(xlsx_file)[0]
            if file_name not in dirs:
                results.append(xlsx_file)
            else:
                dirs.remove(file_name)
                print('REMOVED FROM THE LIST: {}'.format(xlsx_file))
    return results


def main():
    current_path = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_path)
    dirs, xlsx_files = find_excel()
    xlsx_files = exclude_done(dirs, xlsx_files)

    if len(xlsx_files) == 0:
        print('XLSX FILE NOT FOUND OR ALREADY DONE')
        exit()
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
            print('STATISTICS DONE: {}'.format(os.path.splitext(xlsx_file)[0]))


if __name__ == '__main__':
    main()
