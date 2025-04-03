# 文件夹统计工具 (Folder Statistics Tool)

这是一个用于统计电脑所有文件夹大小和文件数量的工具，并生成Excel格式的报告。

## 功能特点

- 自动扫描系统所有可用驱动器
- 递归统计文件夹大小和文件数量
- 生成Excel格式的详细报告
- 文件夹路径支持点击直接打开
- 按文件夹大小降序排序
- 友好的进度显示
- 完善的错误处理

## 安装要求

- Python 3.6+
- 依赖包：
  - humanize
  - openpyxl

## 安装步骤

1. 克隆仓库：
```bash
git clone [仓库地址]
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

1. 以管理员身份运行 PowerShell 或命令提示符
2. 导航到程序目录
3. 运行程序：
```bash
python file_statistics.py
```

## 输出说明

程序会在当前目录生成 `folder_statistics.xlsx` 文件，包含以下信息：
- 文件夹层级
- 文件夹名称
- 完整路径（可点击）
- 文件夹大小
- 文件数量

## 注意事项

1. 建议使用管理员权限运行，以确保能访问所有文件夹
2. 扫描时间取决于文件系统大小，请耐心等待
3. 如果Excel文件被占用，程序会自动使用新文件名保存

## 作者

[您的名字]

## 许可证

私有软件，保留所有权利 