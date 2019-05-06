# 字幕组工作统计脚本

快速统计每个人的工作量与奶茶钱

## 上手指南

以下指南将帮助你在本地机器上安装和运行该项目

### 安装要求

1. Python3.7
2. [pip](https://pypi.org/project/pip/)
3. git

### 安装步骤

```
git clone https://github.com/BurgerTown/subtitile_group_statics.git
pip install -r requirements.txt
```

### 运行

```
python statics.py
```

### 注意事项

1. xlsx文件的格式参考example.xlsx
2. config.py里设置预设参数

## Change Log

v0.1.1 (2019/5/6 17:40)
- 修复在计算时人员的定位问题
- 增加校对增益统计

v0.1.0 (2019/5/6 16:00)
- 发布

## Todo

- csv导出功能
- 自动查找各标签对应行数