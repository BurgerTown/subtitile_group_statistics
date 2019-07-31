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

## Changelog

[更新日志](/CHANGELOG.md)

## Todo

- [ ] 命令行工具(maybe)
- [ ] 输入CSV
- [x] 更新增益算法
- [x] CSV输出功能
- [x] 自动查找各标签对应行数
