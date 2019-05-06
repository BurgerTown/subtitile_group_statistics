# 轴:4 校对:6+2(n-1) 校对:30 后期:33 压制:35
VIDEO_TIME_COL = 3  # 视频时间
TIIMELINE_COL = [4]  # 时间轴
PROOFREAD_COL = [30]  # 校对
TRANSLATION_COLS = [6, 12, 18, 24]  # 翻译
OTHERS_COLS = [33, 35]  # 后期与压制
RELATED_COLS = []

IGNORE_NAMES = ['负责人', '时间轴', '校对', '后期', '压制', '翻译1', '翻译2', '翻译3', '翻译4']

TIMELINE_SALARY = 3
TRANSLATE_SALARY = 15
PROOFREAD_SALARY = 10
SUBTITLE_EDIT_SALARY = 40
COMPRESSION_SALARY = 15

TAGS = ['ID', '时间轴', '翻译', '校对', '后期', '压制', '总计', '总奶茶', '打轴获得奶茶',
        '翻译获得奶茶', '校对获得奶茶', '校对增益奶茶', '总打轴视频时间', '总翻译视频时间', '总校对视频时间']
TAGS_LENGTH = len(TAGS)


RELATED_COLS.extend(TIIMELINE_COL)
RELATED_COLS.extend(PROOFREAD_COL)
RELATED_COLS.extend(TRANSLATION_COLS)
RELATED_COLS.extend(OTHERS_COLS)
