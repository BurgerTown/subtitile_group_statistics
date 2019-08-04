from config import *
import os
import csv
import json
import xlrd


class Statistics():

    def __init__(self):
        self.ASSAULT_COL = 0  # 是否突击
        self.VIDEO_TIME_COL = 0  # 视频时间
        self.TIIMELINE_COL = []  # 时间轴
        self.PROOFREAD_COL = []  # 校对
        self.TRANSLATE_COLS = []  # 翻译
        self.EDIT_COL = []  # 后期
        self.COMPRESSION_COL = []  # 压制
        self.RELATED_COLS = []
        self.IGNORE_NAMES = ['负责人']

        self.participants = {}
        self.statistics = {}
        self.total = {}
        self.file_name = ''

    def read_excel(self, path):
        return xlrd.open_workbook(path)

    def find_RELATED_COLS(self):
        '''
        寻找有关联的关键字来定位
        '''
        row_values, col_indexs = self.sheet.row_values(INFOMATION_ROW), []
        for i in range(len(row_values)):
            if row_values[i] == '视频时长':
                if self.sheet.cell(2, i).ctype == 3:
                    self.VIDEO_TIME_COL = i
            if row_values[i] == '是否突击':
                self.ASSAULT_COL = i
            if row_values[i] == '负责人':
                col_indexs.append(i)
        for col_index in col_indexs:
            row_values = self.sheet.row_values(0)
            if row_values[col_index] == '时间轴':
                self.TIIMELINE_COL.append(col_index)
            if '翻译' in row_values[col_index]:
                self.TRANSLATE_COLS.append(col_index)
            if row_values[col_index] == '校对':
                self.PROOFREAD_COL.append(col_index)
            if row_values[col_index] == '后期':
                self.EDIT_COL.append(col_index)
            if row_values[col_index] == '压制':
                self.COMPRESSION_COL.append(col_index)

        self.RELATED_COLS.extend(self.TIIMELINE_COL)
        self.RELATED_COLS.extend(self.TRANSLATE_COLS)
        self.RELATED_COLS.extend(self.PROOFREAD_COL)
        self.RELATED_COLS.extend(self.EDIT_COL)
        self.RELATED_COLS.extend(self.COMPRESSION_COL)

    def init_dict(self, name):
        '''
        初始化名字字典
        '''
        name_dict = {}
        for tag in TAGS:
            if tag == 'ID':
                name_dict[tag] = name
            else:
                name_dict[tag] = 0
        self.statistics[name] = name_dict

    def has_name(self, name):
        if name not in self.statistics.keys():
            self.init_dict(name)

    def count_row(self, row_value):
        extra_multipier = 1.00
        row_name = []
        if self.sheet.cell(row_value, self.VIDEO_TIME_COL).ctype == 3:
            video_time = xlrd.xldate_as_tuple(self.sheet.cell_value(
                row_value, self.VIDEO_TIME_COL), self.workbook.datemode)
        if self.sheet.cell(row_value, self.ASSAULT_COL).value:
            extra_multipier = 1.00 + ASSAULT_EXTRA

        for col_value in self.RELATED_COLS:
            name = self.sheet.cell(row_value, col_value).value
            if not name:
                continue
            self.has_name(name)
            if name not in row_name:
                row_name.append(name)
                if extra_multipier != 1.00:
                    self.statistics[name]['突击次数'] += 1
            if col_value in self.TIIMELINE_COL:
                self.cal_total_time_related_salary(
                    name, video_time, '时间轴', 0, extra_multipier)

            if col_value in self.TRANSLATE_COLS:
                if self.sheet.cell(row_value, col_value+1).ctype == 3:
                    start_time = xlrd.xldate_as_tuple(self.sheet.cell_value(
                        row_value, col_value+1), self.workbook.datemode)
                if self.sheet.cell(row_value, col_value+2).ctype == 3:
                    end_time = xlrd.xldate_as_tuple(self.sheet.cell_value(
                        row_value, col_value+2), self.workbook.datemode)
                translate_rated = self.sheet.cell(row_value, col_value+4).value
                self.cal_translate_salary(
                    name, start_time, end_time, translate_rated, extra_multipier)

            elif col_value in self.PROOFREAD_COL:
                salary_plus = self.sheet.cell(
                    row_value, self.PROOFREAD_COL[0]+1).value
                self.cal_total_time_related_salary(
                    name, video_time, '校对', salary_plus, extra_multipier)

            elif col_value in self.EDIT_COL:
                self.cal_others_salary(name, '后期', extra_multipier)

            elif col_value in self.COMPRESSION_COL:
                self.cal_others_salary(name, '压制', extra_multipier)

    def begin_collect(self):
        row_number = INFOMATION_ROW + 1
        col_values = self.sheet.col_values(0)
        for i in range(row_number, len(col_values)):
            if col_values[i]:
                self.count_row(row_number)
                row_number += 1

    def cal_total_time_related_salary(self, name, video_time, mode, salary_plus, extra_multipier):
        seconds = video_time[4]*60+video_time[5]
        if mode == '时间轴':
            salary_multiplier = TIMELINE_SALARY
            self.statistics[name]['总打轴视频时间'] += seconds
        else:
            salary_multiplier = PROOFREAD_SALARY
            self.statistics[name]['总校对视频时间'] += seconds

        salary = seconds/60 * salary_multiplier * extra_multipier + salary_plus
        if mode == '时间轴':
            self.statistics[name]['打轴获得奶茶'] += salary
        else:
            self.statistics[name]['校对获得奶茶'] += salary
            self.statistics[name]['校对增益奶茶'] += salary_plus

        self.statistics[name][mode] += 1
        self.statistics[name]['总参与次数'] += 1
        self.statistics[name]['总奶茶'] += salary

    def cal_translate_salary(self, name, start_time, end_time, translate_rated, extra_multipier):
        work_time = end_time[4]*60+end_time[5] - \
            (start_time[4]*60+start_time[5])
        # 时间*(基础工资+打分)/60
        salary = work_time*(TRANSLATE_SALARY+translate_rated) / \
            60 * extra_multipier

        self.statistics[name]['总翻译视频时间'] += work_time
        self.statistics[name]['翻译获得奶茶'] += salary
        self.statistics[name]['翻译'] += 1
        self.statistics[name]['总参与次数'] += 1
        self.statistics[name]['总奶茶'] += salary

    def cal_others_salary(self, name, mode, extra_multipier):
        if mode == '后期':
            self.statistics[name]['后期'] += 1
            salary = SUBTITLE_EDIT_SALARY
        else:
            self.statistics[name]['压制'] += 1
            salary = COMPRESSION_SALARY
        salary *= extra_multipier

        self.statistics[name]['总参与次数'] += 1
        self.statistics[name]['总奶茶'] += salary

    def add_extra(self):
        for key in EXTRA_FIXED.keys():
            self.init_dict(key)
            self.statistics[key]['总奶茶'] = EXTRA_FIXED[key]

    def cal_total(self):
        self.init_dict('总计')
        for tag in TAGS:
            temp = 0
            for name in self.statistics.keys():
                if tag != 'ID':
                    temp += self.statistics[name][tag]
            self.statistics['总计'][tag] = temp

    def beautifier(self):
        for name in self.statistics.keys():
            for key in self.statistics[name].keys():
                if key == 'ID':
                    continue
                data = self.statistics[name][key]
                if '时间' in key and key != '时间轴':
                    if data > 3600:
                        self.statistics[name][key] = f'{data//3600}:{(data - 3600)//60}:{data % 60}'
                    else:
                        self.statistics[name][key] = f'{data//3600}:{data//60}:{data % 60}'
                else:
                    self.statistics[name][key] = round(data, 2)

    def cal_pure_salary(self):
        pure_salary = {}
        for name in self.statistics.keys():
            pure_salary[name] = self.statistics[name]['总奶茶']
        return pure_salary

    def set_env(self):
        if os.path.exists(self.file_name):
            os.chdir(self.file_name)
        else:
            os.mkdir(self.file_name)
            os.chdir(self.file_name)

    def output_json(self):
        with open('{}_pure_salary.json'.format(self.file_name), 'w', encoding='utf8') as f:
            json.dump(self.cal_pure_salary(), f, indent=1, ensure_ascii=False)

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
        self.begin_collect()
        self.add_extra()
        self.cal_total()
        self.beautifier()

        self.set_env()
        self.output_json()
        self.output_csv()
