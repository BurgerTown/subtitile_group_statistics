from config import *
import os
import json
import xlrd


def find_excel():
    file_list = os.listdir()
    for file_name in file_list:
        if os.path.isfile(file_name):
            file_name = file_name.split('.')
            if file_name[len(file_name)-1] == 'xlsx':
                name = ''
                for i in range(len(file_name)-2):
                    name += file_name[i]
                return name
    return ''


def read_excel(path):
    return xlrd.open_workbook(path)


def collect_participants(sheet):
    participants = {}
    for i in RELATED_COLS:
        names = sheet.col_values(i)
        for name in names:
            if name != '' and name not in IGNORE_NAMES:
                if name not in participants.keys():
                    participants[name] = {}
                    participants[name]['总计'] = 1
                    if '翻译' in names[0]:
                        participants[name]['翻译'] = 1
                    else:
                        participants[name][names[0]] = 1
                else:
                    participants[name]['总计'] += 1
                    if '翻译' in names[0]:
                        if '翻译' in participants[name].keys():
                            participants[name]['翻译'] += 1
                        else:
                            participants[name]['翻译'] = 1
                    else:
                        if names[0] in participants[name].keys():
                            participants[name][names[0]] += 1
                        else:
                            participants[name][names[0]] = 1
    return participants


def cal_time_related_salary(workbook, sheet, name, mode):
    salary = 0
    seconds = 0
    if mode == '时间轴':
        salary_multiplier = TIMELINE_SALARY
        col = TIIMELINE_COL[0]
    else:
        salary_multiplier = PROOFREAD_SALARY
        col = PROOFREAD_COL[0]
    col_values = sheet.col_values(col)  # 获取对应列的数据
    for col_value in col_values:  # 遍历对应列的数据
        if col_value == name:  # 如果与传入的名字匹配
            row_index = col_values.index(col_value)  # 获取其索引值 作为行的值
            col_index = VIDEO_TIME_COL
            if sheet.cell(row_index, col_index).ctype == 3:  # 判断目标单元格的数据类型是否时间
                video_time = xlrd.xldate_as_tuple(sheet.cell_value(
                    row_index, col_index), workbook.datemode)
                salary_plus = 0
                if mode == '校对':
                    salary_plus = sheet.cell(row_index, col_index+1).value
                temp_time = video_time[4]*60+video_time[5]
                seconds += temp_time
                salary += temp_time*salary_multiplier/60+salary_plus
    return salary, seconds


def cal_translate_salary(workbook, sheet, name):
    salary = 0
    seconds = 0
    for translation_col in TRANSLATION_COLS:
        col_values = sheet.col_values(translation_col)
        for col_value in col_values:
            if col_value == name:
                temp_time = 0
                row_index = col_values.index(col_value)
                col_index = translation_col
                if sheet.cell(row_index, col_index+1).ctype == 3:
                    start_time = xlrd.xldate_as_tuple(sheet.cell_value(
                        row_index, col_index+1), workbook.datemode)
                if sheet.cell(row_index, col_index+2).ctype == 3:
                    end_time = xlrd.xldate_as_tuple(sheet.cell_value(
                        row_index, col_index+2), workbook.datemode)
                salary_multiplier = sheet.cell(row_index, col_index+4).value
                temp_time = end_time[4] * 60+end_time[5] - \
                    start_time[4]*60-start_time[5]
                seconds += temp_time
                salary += temp_time*salary_multiplier*TRANSLATE_SALARY/60
    return salary, seconds


def cal_others_salary(participant):
    salary = 0
    if '后期' in participant.keys():
        salary += participant['后期'] * SUBTITLE_EDIT_SALARY
    if '压制' in participant.keys():
        salary += participant['压制'] * COMPRESSION_SALARY
    return salary


def cal_time_and_salary(participants, workbook, sheet):
    for name in list(participants.keys()):
        participant = participants[name]
        salary = 0
        participant['奶茶'] = {}
        if '时间轴' in participant.keys():
            timeline_salary = 0
            timeline_time = 0
            timeline_salary, timeline_time = cal_time_related_salary(
                workbook, sheet, name, '时间轴')
            salary += timeline_salary
            participant['总打轴视频时间'] = timeline_time
            participant['奶茶']['打轴获得奶茶'] = timeline_salary
        if '翻译' in participant.keys():
            translate_time = 0
            translate_time = 0
            translate_salary, translate_time = cal_translate_salary(
                workbook, sheet, name)
            salary += translate_salary
            participant['总翻译视频时间'] = translate_time
            participant['奶茶']['翻译获得奶茶'] = translate_salary
        if '校对' in participant.keys():
            proofread_salary = 0
            proofread_time = 0
            proofread_salary, proofread_time = cal_time_related_salary(
                workbook, sheet, name, '校对')
            salary += proofread_salary
            participant['总校对视频时间'] = proofread_time
            participant['奶茶']['校对获得奶茶'] = proofread_salary
        if '后期' or '压制' in participant.keys():
            salary += cal_others_salary(participant)
        participant['奶茶']['总奶茶'] = int(salary)
    return participants


def cal_pure_salary(statics):
    pure_salary = {}
    for name in statics.keys():
        pure_salary[name] = statics[name]['奶茶']['总奶茶']
    return pure_salary


def output_csv(statics):
    pass


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    xlsx_file = find_excel()
    if xlsx_file == '':
        print('XLSX FILE NOT FOUND')
        exit()
    else:
        workbook = read_excel('{}.xlsx'.format(xlsx_file))
        sheet = workbook.sheet_by_index(0)
        participants = collect_participants(sheet)
        statics = cal_time_and_salary(participants, workbook, sheet)
        with open('{}.json'.format(xlsx_file), 'w', encoding='utf8') as f:
            json.dump(statics, f, indent=1, ensure_ascii=False)
        with open('{}_pure_salary.json'.format(xlsx_file), 'w', encoding='utf8') as f:
            json.dump(cal_pure_salary(statics), f,
                      indent=1, ensure_ascii=False)


if __name__ == '__main__':
    main()
