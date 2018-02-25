## 使用 js-xlsx 解析excel文件

### 读取文件
```js
import XLSX from 'xlsx';
const workbook = XLSX.readFile('someExcel.xlsx', opts);
```

### 读取表

```js
// 获取 Excel 中所有表名
const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
// 根据表名获取对应某张表
const worksheet = workbook.Sheets[sheetNames[0]];
```

### 通过 worksheet[address] 来操作表格，以 ! 开头的 key 是特殊的字段。

```js
// 获取 A1 单元格对象
let a1 = worksheet['A1']; // 返回 { v: 'hello', t: 's', ... }
// 获取 A1 中的值
a1.v // 返回 'hello'

// 获取表的有效范围
worksheet['!ref'] // 返回 'A1:B20'
worksheet['!range'] // 返回 range 对象，{ s: { r: 0, c: 0}, e: { r: 100, c: 2 } }

// 获取合并过的单元格
worksheet['!merges'] // 返回一个包含 range 对象的列表，[ {s: { r: 0, c: 0 }, c: { r: 2, c: 1 } } ]
```

### 转化为json


```js
XLSX.utils.sheet_to_json(worksheet)
```