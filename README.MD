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
python main.py
```

### 注意事项

1. xlsx文件的格式参考example.xlsx
2. config.py里设置预设参数

## Change Log

v0.1.8 (2019/7/5 11:30)

- 修复`example.xlsx`一个小错误
- `example.xlsx` -> `Example.xlsx`
- 增加重复检测功能
- 增加命令行输出

v0.1.7 (2019/6/3 23:00)

- 结果会输出到相对应文件名的文件夹
- 更新`example.xlsx`
- 优化代码

v0.1.6 (2019/5/10 13:20)

- 改为对象写法
- 优化代码
- 更新README.MD的`运行`与`Todo`

v0.1.5 (2019/5/8 11:40)

- 输出CSV时 所有小数保留两位 时间使用`hh:mm:ss`格式表示

v0.1.4 (2019/5/7 17:50)

- 更新增益算法
- 更新excel文件中的打分部分

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

- [ ] 命令行工具(maybe)
- [ ] 输入CSV
- [x] 更新增益算法
- [x] CSV输出功能
- [x] 自动查找各标签对应行数
