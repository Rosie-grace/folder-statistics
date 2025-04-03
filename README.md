# 文件夹统计工具 (Folder Statistics Tool)

这是一个用于统计电脑所有文件夹大小和文件数量的工具，并生成Excel格式的报告。支持扫描单个驱动器或整个系统的所有驱动器。

## 功能特点

- 自动扫描系统所有可用驱动器
- 递归统计文件夹大小和文件数量
- 生成Excel格式的详细报告
- 文件夹路径支持点击直接打开
- 按文件夹大小降序排序
- 友好的进度显示
- 完善的错误处理
- 支持大文件系统扫描
- 自动处理文件访问权限问题
- Excel报告自动列宽调整

## 安装要求

- Python 3.6+
- 依赖包：
  - humanize：用于文件大小的人性化显示
  - openpyxl：用于生成Excel报告

## 安装步骤

1. 克隆仓库：
```bash
git clone https://github.com/Rosie-grace/folder-statistics.git
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

1. 扫描特定驱动器：
```bash
python file_statistics.py E:
```

2. 扫描所有驱动器：
```bash
python file_statistics.py
```

注意：建议以管理员身份运行 PowerShell 或命令提示符，以确保能访问所有文件夹。

## 输出说明

程序会在当前目录生成 `folder_statistics.xlsx` 文件，包含以下信息：
- 文件夹层级：显示文件夹的嵌套深度
- 文件夹名称：当前文件夹的名称
- 完整路径：可点击直接打开文件夹
- 文件夹大小：人性化显示（KB/MB/GB）
- 文件数量：文件夹中的文件总数

## 报告示例

生成的Excel报告格式如下：

| 层级 | 文件夹名称 | 完整路径 | 大小 | 文件数量 |
|------|------------|----------|------|-----------|
| 0    | Windows    | C:\Windows | 25.6 GB | 1250 |
| 1    | System32   | C:\Windows\System32 | 15.2 GB | 850 |
| ...  | ...        | ...      | ...  | ... |

## 注意事项

1. 权限要求：
   - 建议使用管理员权限运行
   - 对于无权限访问的文件夹会自动跳过
   - 错误信息会在控制台显示

2. 性能考虑：
   - 扫描时间取决于文件系统大小
   - 建议先扫描单个驱动器测试
   - 大型系统扫描可能需要较长时间

3. 文件处理：
   - 如果Excel文件被占用，会自动使用新文件名
   - 支持超大文件夹的统计
   - 自动处理特殊字符和长路径

## 常见问题

1. 访问被拒绝
   - 以管理员身份运行程序
   - 检查文件夹权限设置

2. 报告打不开
   - 确保已安装Excel或WPS
   - 检查文件是否被其他程序占用
   - 尝试使用新文件名保存

3. 扫描速度慢
   - 属于正常现象，取决于文件系统大小
   - 可以先扫描单个驱动器
   - 关闭其他占用磁盘的程序

## 作者

[Rosie-grace](https://github.com/Rosie-grace)

## 许可证

私有软件，保留所有权利

## 更新日志

### v1.0.0 (2025-04-03)
- 初始版本发布
- 支持多驱动器扫描
- Excel格式报告生成 