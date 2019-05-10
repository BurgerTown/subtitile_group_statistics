from config import *
import os
import csv
import json
import xlrd


class Statistics():

    def __init__(self):
        self.VIDEO_TIME_COL = []  # 视频时间
        self.TIIMELINE_COL = []  # 时间轴
        self.PROOFREAD_COL = []  # 校对
        self.TRANSLATE_COLS = []  # 翻译
        self.OTHERS_COLS = []  # 后期与压制
        self.RELATED_COLS = []
        self.IGNORE_NAMES = ['负责人']

        self.participants = {}
        self.statistics = {}
        self.total = {}
        self.file_name = ''

    def read_excel(self, path):
        return xlrd.open_workbook(path)

    def find_IGNORE_NAMES(self):
        row_values = self.sheet.row_values(0)
        for row_value in row_values:
            if row_value != '':
                self.IGNORE_NAMES.append(row_value)

    def find_RELATED_COLS(self):
        row_values, col_indexs = self.sheet.row_values(1), []
        for i in range(len(row_values)):
            if row_values[i] == '负责人':
                col_indexs.append(i)
            if row_values[i] == '视频时长':
                if self.sheet.cell(2, i).ctype == 3:
                    self.VIDEO_TIME_COL.append(i)
        for col_index in col_indexs:
            row_values = self.sheet.row_values(0)
            if row_values[col_index] == '时间轴':
                self.TIIMELINE_COL.append(col_index)
            if '翻译' in row_values[col_index]:
                self.TRANSLATE_COLS.append(col_index)
            if row_values[col_index] == '校对':
                self.PROOFREAD_COL.append(col_index)
            if row_values[col_index] == '后期':
                self.OTHERS_COLS.append(col_index)
            if row_values[col_index] == '压制':
                self.OTHERS_COLS.append(col_index)
        self.RELATED_COLS.extend(self.TIIMELINE_COL)
        self.RELATED_COLS.extend(self.PROOFREAD_COL)
        self.RELATED_COLS.extend(self.TRANSLATE_COLS)
        self.RELATED_COLS.extend(self.OTHERS_COLS)

    def init_dict(self, name_dict):
        for tag in TAGS:
            name_dict[tag] = 0
        return name_dict

    def collect_participants(self):
        participants = {}
        for i in self.RELATED_COLS:
            names = self.sheet.col_values(i)
            for name in names:
                if name != '' and name not in self.IGNORE_NAMES:
                    if name not in participants.keys():
                        participants[name] = {}
                        participants[name] = self.init_dict(participants[name])
                        participants[name]['参与次数'] = 1
                        if '翻译' in names[0]:
                            participants[name]['翻译'] = 1
                        else:
                            participants[name][names[0]] = 1
                    else:
                        participants[name]['参与次数'] += 1
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
        self.participants = participants

    def cal_time_related_salary(self, name, mode):
        salary = seconds = total_salary_plus = 0
        if mode == '时间轴':
            salary_multiplier = TIMELINE_SALARY
            col = self.TIIMELINE_COL[0]
        else:
            salary_multiplier = PROOFREAD_SALARY
            col = self.PROOFREAD_COL[0]
        col_values = self.sheet.col_values(col)  # 获取对应列的数据
        for i in range(len(col_values)):  # 遍历对应列的数据
            if col_values[i] == name:  # 如果与传入的名字匹配
                row_index, col_index = i, self.VIDEO_TIME_COL[0]
                # 判断目标单元格的数据类型是否时间
                if self.sheet.cell(row_index, col_index).ctype == 3:
                    video_time = xlrd.xldate_as_tuple(self.sheet.cell_value(
                        row_index, col_index), self.workbook.datemode)
                    salary_plus = 0
                    if mode == '校对':
                        salary_plus = self.sheet.cell(
                            row_index, self.PROOFREAD_COL[0]+1).value
                    temp_time = video_time[4]*60+video_time[5]
                    seconds += temp_time
                    # 时间*工资/60+校对增益
                    salary += temp_time*salary_multiplier/60+salary_plus
                    total_salary_plus += salary_plus
        return salary, seconds, total_salary_plus

    def cal_translate_salary(self, name):
        salary = seconds = 0
        for translation_col in self.TRANSLATE_COLS:
            col_values = self.sheet.col_values(translation_col)
            for i in range(len(col_values)):
                if col_values[i] == name:
                    temp_time, row_index, col_index = 0, i, translation_col
                    if self.sheet.cell(row_index, col_index+1).ctype == 3:
                        start_time = xlrd.xldate_as_tuple(self.sheet.cell_value(
                            row_index, col_index+1), self.workbook.datemode)
                    if self.sheet.cell(row_index, col_index+2).ctype == 3:
                        end_time = xlrd.xldate_as_tuple(self.sheet.cell_value(
                            row_index, col_index+2), self.workbook.datemode)
                    translate_rated = self.sheet.cell(
                        row_index, col_index+4).value
                    temp_time = end_time[4]*60+end_time[5] - \
                        (start_time[4]*60+start_time[5])
                    seconds += temp_time
                    # 时间*(基础工资+打分)/60
                    salary += temp_time*(TRANSLATE_SALARY+translate_rated)/60
        return salary, seconds

    def cal_others_salary(self, participant):
        salary = 0
        if '后期' in participant.keys():
            salary += participant['后期'] * SUBTITLE_EDIT_SALARY
        if '压制' in participant.keys():
            salary += participant['压制'] * COMPRESSION_SALARY
        # 这两个都是 次数*工资
        return salary

    def cal_time_and_salary(self):
        for name in self.participants.keys():
            participant, salary = self.participants[name], 0
            if '时间轴' in participant.keys():
                timeline_salary = timeline_time = total_salary_plus = 0
                timeline_salary, timeline_time, total_salary_plus = self.cal_time_related_salary(
                    name, '时间轴')
                salary += timeline_salary
                participant['总打轴视频时间'] = timeline_time
                participant['打轴获得奶茶'] = timeline_salary
            if '翻译' in participant.keys():
                translate_time = translate_time = 0
                translate_salary, translate_time = self.cal_translate_salary(
                    name)
                salary += translate_salary
                participant['总翻译视频时间'] = translate_time
                participant['翻译获得奶茶'] = translate_salary
            if '校对' in participant.keys():
                proofread_salary = proofread_time = total_salary_plus = 0
                proofread_salary, proofread_time, total_salary_plus = self.cal_time_related_salary(
                    name, '校对')
                salary += proofread_salary
                participant['总校对视频时间'] = proofread_time
                participant['校对获得奶茶'] = proofread_salary
                participant['校对增益奶茶'] = total_salary_plus
            if '后期' or '压制' in participant.keys():
                salary += self.cal_others_salary(participant)
            participant['总奶茶'] = salary
            self.statistics[name] = participant

    def cal_total(self):
        self.total = self.init_dict(self.total)
        for tag in TAGS:
            temp = 0
            for name in self.statistics.keys():
                if tag != 'ID':
                    temp += self.statistics[name][tag]
            self.total[tag] = temp
        self.total['ID'] = '总计'

    def beautifier(self):
        for name in self.statistics.keys():
            for key in self.statistics[name].keys():
                data = self.statistics[name][key]
                if '时间' in key and key != '时间轴':
                    self.statistics[name][key] = '{}:{}:{}'.format(
                        data//3600, data//60, data % 60)
                else:
                    self.statistics[name][key] = round(data, 2)
        for key in self.total.keys():
            data = self.total[key]
            if key == 'ID':
                continue
            elif '时间' in key and key != '时间轴':
                # dirty...
                if data > 3600:
                    self.total[key] = '{}:{}:{}'.format(
                        data//3600, (data-3600)//60, data % 60)
                else:
                    self.total[key] = '{}:{}:{}'.format(
                        data//3600, data//60, data % 60)
            else:
                self.total[key] = round(data, 2)

    def cal_pure_salary(self):
        pure_salary = {}
        for name in self.statistics.keys():
            pure_salary[name] = self.statistics[name]['总奶茶']
        return pure_salary

    def output_csv(self):
        with open('{}.csv'.format(self.file_name), 'w', encoding='utf_8_sig', newline='') as f:
            f_csv = csv.DictWriter(f, TAGS)
            f_csv.writeheader()
            for name in self.statistics.keys():
                self.statistics[name]['ID'] = name
                f_csv.writerow(self.statistics[name])
            f_csv.writerow(self.total)

    def count(self, xlsx_file):
        self.file_name = os.path.splitext(xlsx_file)[0]
        self.workbook = self.read_excel(xlsx_file)
        self.sheet = self.workbook.sheet_by_index(0)

        self.find_RELATED_COLS()
        self.find_IGNORE_NAMES()
        self.collect_participants()
        self.cal_time_and_salary()
        self.cal_total()
        self.beautifier()

        with open('{}_pure_salary.json'.format(self.file_name), 'w', encoding='utf8') as f:
            json.dump(self.cal_pure_salary(), f, indent=1, ensure_ascii=False)
        self.output_csv()
