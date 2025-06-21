# 旅程记录生成器 (Trip Generator)

一个智能的旅程记录生成工具，能够自动生成符合特定要求的旅程报告并导出为 Excel 文件。

## 🆕 最新更新

- ✅ **Google Maps API 集成**: 支持真实距离计算
- ✅ **荷兰语界面**: Excel 表头和输出信息使用荷兰语
- ✅ **扩展城市列表**: 包含荷兰村庄 (dorpen) 和比利时弗拉芒区域
- ✅ **环境变量支持**: 安全便捷的 API 密钥管理

## 功能特点

### 🎯 智能距离分布

- **40%** 短途行程 (< 100 公里)
- **40%** 中途行程 (100-300 公里)
- **20%** 长途行程 (> 300 公里)

### 🗺️ 真实距离计算

- 集成 Google Maps API
- 基于实际驾驶路线计算距离
- 自动避开收费路段
- 提供详细的行驶时间信息

### 🏷️ 智能地区识别

- 自动识别荷兰和比利时城市
- 比利时城市自动添加"BE"标识
- 支持多种地址格式匹配

### 🔄 重复限制控制

- **目的地重复限制**: 每个目的地最多使用 3 次
- **同一天行程限制**: 每天最多安排 2 次行程
- 智能回退机制，自动寻找替代方案

### 📊 详细统计报告

- 实时显示生成进度
- 目的地使用统计
- 日期分布统计
- 距离分布分析

## 安装要求

```bash
pip install pandas openpyxl googlemaps
```

## 环境配置

### 方法 1: 使用.env 文件 (推荐)

创建`.env`文件：

```
GOOGLE_MAPS_API_KEY=your_api_key_here
```

### 方法 2: 设置环境变量

```bash
export GOOGLE_MAPS_API_KEY='your_api_key_here'
```

### 方法 3: 命令行参数

```bash
python trip_generator.py --google-api-key YOUR_API_KEY [其他参数]
```

## 使用方法

### 基本用法

```bash
python trip_generator.py --year 2025 --quarter 1 --target-km 5000 --address "您的起始地址"
```

### 完整参数示例

```bash
python trip_generator.py \
  --year 2025 \
  --quarter 1 \
  --target-km 5000 \
  --address "您的起始地址" \
  --output 我的旅程报告.xlsx \
  --google-api-key YOUR_API_KEY \
  --json \
  --seed 12345
```

## 参数说明

| 参数               | 类型 | 必需 | 说明                      |
| ------------------ | ---- | ---- | ------------------------- |
| `--year`           | int  | ✅   | 年份 (例: 2025)           |
| `--quarter`        | int  | ✅   | 季度 (1, 2, 3, 4)         |
| `--target-km`      | int  | ✅   | 目标总公里数              |
| `--address`        | str  | ✅   | 起始地址                  |
| `--output`         | str  | ❌   | 输出 Excel 文件名         |
| `--google-api-key` | str  | ❌   | Google Maps API 密钥      |
| `--json`           | flag | ❌   | 同时生成 JSON 文件        |
| `--seed`           | int  | ❌   | 随机种子 (用于可重现结果) |

## 输出文件

### Excel 文件格式

- **Datum**: 日期 (DD-MM-YYYY)
- **Bestemming**: 目的地
- **Omschrijving**: 描述 (默认: "klant bezoeken")
- **Totale afstand (km)**: 总距离 (往返)

### 统计信息

- 总行程数和总公里数
- 距离分布统计
- 目的地使用频率
- 日期分布情况

## 智能限制机制

### 目的地重复限制 (最多 3 次)

```
✨ 添加行程: Delft (20km - 短途)
✨ 添加行程: Delft (2/3)(20km - 短途)
✨ 添加行程: Delft (3/3)(20km - 短途)
⚠️  Delft已达到3次使用限制，将自动选择其他目的地
```

### 同一天行程限制 (最多 2 次)

```
✨ 添加行程: Utrecht (137km - 中途)
✨ 添加行程: Leiden (48km - 短途) [2/2日程]
⚠️  该日期已有2次行程，将自动选择其他日期
```

### 距离分布控制

系统会智能选择目的地以达到目标分布：

- 当短途行程不足时，优先选择附近城市
- 当长途行程不足时，优先选择远距离或比利时城市
- 实时显示各类型距离的完成进度

## 支持的地区

### 荷兰城市

- 主要城市: Amsterdam, Rotterdam, Den Haag, Utrecht 等
- 中小城市: Delft, Leiden, Gouda, Alphen aan den Rijn 等
- 村庄: Volendam, Marken, Edam 等

### 比利时弗拉芒区 (自动添加 BE 标识)

- 主要城市: Antwerpen BE, Gent BE, Brugge BE 等
- 中小城市: Kortrijk BE, Hasselt BE, Leuven BE 等

## 故障排除

### API 密钥问题

```
❌ 未找到Google Maps API密钥!
   请使用以下方式之一:
   1. 设置环境变量: export GOOGLE_MAPS_API_KEY='your_api_key'
   2. 使用命令行参数: --google-api-key YOUR_API_KEY
   3. 确保.env文件存在且已加载
```

### 网络连接问题

```
⚠️  API调用异常: 城市名 - 网络错误
⏭️  跳过目的地: 城市名
```

### 限制达到问题

```
⚠️  所有目的地都已达到3次使用限制或无法访问
⚠️  无法找到合适的日期（所有日期都已有2次行程）
```

## 示例输出

### 控制台输出

```
🚗 Genereren van reisverslag voor 2025 Q1...
📍 Startlocatie: 您的起始地址
🎯 Doel kilometers: 5000
✅ Google Maps API已连接

🎯 距离分布目标:
   短途 (<100km): 2000km (40%)
   中途 (100-300km): 2000km (40%)
   长途 (>300km): 1000km (20%)

✨ 添加行程: Leiden (48km - 短途)
✨ 添加行程: Nijmegen (274km - 中途)
✨ 添加行程: Antwerpen BE (364km - 长途)

📊 最终距离分布:
   短途 (<100km): 1987km (39.7%)
   中途 (100-300km): 2089km (41.8%)
   长途 (>300km): 924km (18.5%)
   总计: 5000km

✅ Reisverslag geëxporteerd naar: reisverslag_2025_Q1.xlsx
```

### Excel 文件内容

| Datum      | Bestemming   | Omschrijving   | Totale afstand (km) |
| ---------- | ------------ | -------------- | ------------------- |
| 04-01-2025 | Leiden       | klant bezoeken | 48                  |
| 15-01-2025 | Nijmegen     | klant bezoeken | 274                 |
| 22-01-2025 | Antwerpen BE | klant bezoeken | 364                 |
| ...        | ...          | ...            | ...                 |
| Totaal     | 25 ritten    | 2025 Q1        | 5000                |

## 版本历史

- **v1.0**: 基础功能实现
- **v1.1**: 添加 Google Maps API 集成
- **v1.2**: 实现距离分布控制
- **v1.3**: 添加比利时城市支持
- **v1.4**: 实现目的地重复限制
- **v1.5**: 添加同一天行程限制
- **v2.0**: 重构命令行界面，使用命名参数

## 许可证

本项目仅供个人使用。请确保遵守 Google Maps API 的使用条款。
