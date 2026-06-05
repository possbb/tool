# 时间列自动转行工具

这个工具提供两种使用方式：网页版和命令行版。

## 网页版使用方法

1. 双击 `启动网页版.bat`
2. 浏览器会打开 `http://127.0.0.1:5000`
3. 上传表格文件
4. 点击“开始转换并下载”

网页会自动识别表格中代表时间的列，并输出为“时间 / 值”格式的长表。

## 支持格式

- 输入：`.xlsx`、`.xls`、`.xlsm`、`.csv`、`.tsv`
- 输出：`.xlsx`

## 命令行使用方法

如果想直接命令行运行，可以先安装依赖：

```powershell
pip install -r requirements.txt
```

然后运行：

```powershell
python .\time_columns_to_rows.py "你的表格.xlsx"
```

指定输出文件：

```powershell
python .\time_columns_to_rows.py "你的表格.xlsx" -o "转置结果.xlsx"
```

指定 Excel 工作表：

```powershell
python .\time_columns_to_rows.py "你的表格.xlsx" -s "Sheet1"
```

修改输出列名：

```powershell
python .\time_columns_to_rows.py "你的表格.xlsx" --time-name "月份" --value-name "金额"
```

## 可自动识别的时间列名示例

- `2024-01`
- `2024/01/31`
- `2024年1月`
- `1月`
- `2024年`
- `Q1 2024`
- `2024 Q1`
- `Jan 2024`

如果你的表头格式比较特殊，可以继续补充识别规则。
