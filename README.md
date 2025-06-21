# 旅程记录生成器 (Reisverslag Generator)

这是一个 Python 脚本，可以自动生成旅程记录并导出到 Excel 文件。

## 🆕 最新更新

- ✅ **Google Maps API 集成**: 支持真实距离计算
- ✅ **荷兰语界面**: Excel 表头和输出信息使用荷兰语
- ✅ **扩展城市列表**: 包含荷兰村庄 (dorpen) 和比利时弗拉芒区域
- ✅ **环境变量支持**: 安全便捷的 API 密钥管理

## 功能特点

- ✅ 使用 `argparse` 处理命令行参数
- ✅ 支持指定年份、季度、目标公里数和起始位置
- ✅ 自动生成随机的荷兰城市/村庄和比利时弗拉芒区域目的地
- ✅ 生成 DD-MM-YYYY 格式的日期，按季度排列
- ✅ **真实距离计算**: 使用 Google Maps API 或随机生成
- ✅ 导出到 Excel 文件，**荷兰语表头**
- ✅ 可选保存 JSON 格式
- ✅ **环境变量支持**: 安全的 API 密钥配置

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
python trip_generator.py <年份> <季度> <目标公里数> "<起始位置>"
```

### 示例

```bash
# 基本使用（随机距离）
python trip_generator.py 2025 1 5000 "de genestetlaan 299 den haag"

# 使用环境变量中的 Google Maps API 密钥（推荐）
python trip_generator.py 2025 1 5000 "de genestetlaan 299 den haag"

# 直接指定 Google Maps API 密钥
python trip_generator.py 2025 1 5000 "de genestetlaan 299 den haag" --google-api-key YOUR_API_KEY

# 指定输出文件名
python trip_generator.py 2025 1 5000 "Amsterdam Centraal" --output "mijn_reisverslag.xlsx"

# 同时保存JSON文件
python trip_generator.py 2025 1 5000 "Rotterdam Centraal" --json

# 使用随机种子确保结果可重现
python trip_generator.py 2025 1 5000 "Utrecht Centraal" --seed 12345
```

### 参数说明

| 参数               | 类型 | 必需 | 说明                     |
| ------------------ | ---- | ---- | ------------------------ |
| `year`             | int  | 是   | 年份（如 2025）          |
| `quarter`          | int  | 是   | 季度（1, 2, 3, 4）       |
| `target_km`        | int  | 是   | 目标公里数               |
| `start_location`   | str  | 是   | 起始位置                 |
| `--output`, `-o`   | str  | 否   | 输出 Excel 文件名        |
| `--json`           | flag | 否   | 同时保存 JSON 文件       |
| `--seed`           | int  | 否   | 随机种子（用于重现结果） |
| `--google-api-key` | str  | 否   | Google Maps API 密钥     |

## 输出格式

生成的 Excel 文件包含以下**荷兰语**列：

- **Datum**: DD-MM-YYYY 格式
- **Bestemming**: 随机选择的荷兰城市/村庄或比利时弗拉芒区域
- **Omschrijving**: "klant bezoeken"
- **Totale afstand (km)**: 计算的往返距离

## Google Maps API 设置

### 1. 获取 API 密钥

1. 访问 [Google Cloud Console](https://console.cloud.google.com/)
2. 创建新项目或选择现有项目
3. 启用 **Distance Matrix API**
4. 创建 API 密钥
5. （可选）限制 API 密钥使用范围

### 2. 配置环境变量（推荐方式）

#### 🔐 安全最佳实践

⚠️ **重要安全提醒：**

- ❌ **绝对不要**将 API 密钥直接写在代码中
- ❌ **绝对不要**将包含真实 API 密钥的文件提交到版本控制
- ✅ **务必**使用环境变量或配置文件管理敏感信息
- ✅ **确保** `.env` 文件已被 `.gitignore` 忽略

#### 🔧 快速设置

**安全配置步骤：**

1. **复制配置模板**

```bash
cp .env.example .env
```

2. **编辑 .env 文件**

```bash
# 编辑 .env 文件，替换 YOUR_ACTUAL_API_KEY_HERE 为您的真实API密钥
nano .env
```

3. **加载环境变量**

```bash
source .env
```

#### 🛠️ 手动设置

**方法 1: 临时设置（当前会话）**

```bash
export GOOGLE_MAPS_API_KEY="你的API密钥"
```

**方法 2: 永久设置 - Zsh (macOS 默认)**

```bash
echo 'export GOOGLE_MAPS_API_KEY="你的API密钥"' >> ~/.zshrc
source ~/.zshrc
```

**方法 3: 永久设置 - Bash**

```bash
echo 'export GOOGLE_MAPS_API_KEY="你的API密钥"' >> ~/.bashrc
source ~/.bashrc
```

**方法 4: 使用 .env 文件（推荐）**

```bash
# 创建 .env 文件（此文件不会被提交到版本控制）
echo 'GOOGLE_MAPS_API_KEY=你的API密钥' > .env
source .env
```

### 3. 使用 API 密钥

设置环境变量后，直接运行脚本即可：

```bash
# 脚本会自动使用环境变量中的API密钥
python trip_generator.py 2025 1 5000 "Amsterdam Centraal"

# 也可以通过命令行参数覆盖环境变量
python trip_generator.py 2025 1 5000 "Amsterdam Centraal" --google-api-key OTHER_API_KEY
```

### 4. 测试 API 连接

```bash
python google_api_example.py
```

## 城市和地区覆盖

脚本现在包含：

### 🇳🇱 荷兰

- **主要城市**: Amsterdam, Rotterdam, Den Haag, Utrecht 等
- **村庄 (Dorpen)**: Volendam, Marken, Edam, Monnickendam 等
- **涵盖所有省份**: 从 Groningen 到 Limburg

### 🇧🇪 比利时弗拉芒区域 (Vlaanderen)

- **主要城市**: Antwerpen, Gent, Brugge, Leuven 等
- **小城镇**: Knokke-Heist, Blankenberge, Ieper 等
- **历史地区**: 包含西弗拉芒省各地

## 距离计算

### 使用 Google Maps API（推荐）

- ✅ 真实驾驶距离
- ✅ 避免收费道路
- ✅ 考虑实际路况
- ⚠️ 需要有效的 API 密钥

### 随机生成（后备选项）

- 🎲 50-400 公里单程范围
- 🎲 适合测试和演示
- 🆓 无需 API 密钥

## 注意事项

- Excel 表头和界面信息使用荷兰语
- 日期生成在指定季度内随机分布
- 脚本会确保总公里数接近目标值（允许 10%误差）
- 生成的旅程按日期排序
- Google Maps API 调用失败时会自动回退到随机距离
- 环境变量优先级高于命令行参数
- **安全提醒**: `.env` 文件包含敏感信息，已被 `.gitignore` 忽略

## 文件结构

```
trip_generator.py       # 主脚本
requirements.txt        # 依赖列表
google_api_example.py   # Google API 使用示例
.env.example           # 环境变量配置示例（安全提交）
.gitignore             # Git忽略文件（包含.env）
README.md              # 本说明文件
```

⚠️ **注意**: `.env` 文件包含敏感信息，不会出现在版本控制中

## 许可证

MIT License
