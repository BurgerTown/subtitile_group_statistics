# 字幕组工作统计脚本

快速统计每个人的工作量与奶茶钱

## 上手指南

以下指南将帮助你在本地机器上安装和运行该项目

### 安装要求

1. Python3.7
2. [pip](https://pypi.org/project/pip/)
3. git

### 安装步骤

```shell
git clone https://github.com/BurgerTown/subtitile_group_statistics.git
pip install -r requirements.txt
```

### 运行

```shell
python statistics.py
```

### 注意事项

1. xlsx文件的格式参考example.xlsx
2. config.py里设置预设参数

## Change Log

v0.1.3 (2019/5/7 11:00)

- 增加自动查找各标签对应行数的功能 不需要手动设置了
- 增加`总计`统计 结果会输出到CSV
- 更新`config.py`的备注

v0.1.2 (2019/5/6 23:00)

- 修复在计算时的索引问题 此问题会导致漏算最后一行
- 增加CSV输出功能
- 移除了输出`example.json`的功能 由CSV代替

v0.1.1 (2019/5/6 17:40)

- 修复在计算时人员的定位问题
- 增加校对增益统计

v0.1.0 (2019/5/6 16:00)

- 发布

## Todo

- [ ] 更新增益算法
- [x] CSV输出功能
- [x] 自动查找各标签对应行数