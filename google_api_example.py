#!/usr/bin/env python3
"""
Google Maps API使用示例
展示如何获取和使用Google Maps API密钥

要使用Google Maps API，请按以下步骤操作：

1. 访问 Google Cloud Console: https://console.cloud.google.com/
2. 创建一个新项目或选择现有项目
3. 启用 Distance Matrix API
4. 创建API密钥
5. （可选）限制API密钥的使用范围以提高安全性

API密钥配置方法：
"""

import os

# 使用Google Maps API的示例命令：
example_commands = [
    # 使用环境变量（推荐）
    'export GOOGLE_MAPS_API_KEY="your_actual_api_key"',
    'python trip_generator.py 2025 1 5000 "de genestetlaan 299 den haag"',
    '',
    # 使用命令行参数
    'python trip_generator.py 2025 1 5000 "de genestetlaan 299 den haag" --google-api-key YOUR_API_KEY',
    '',
    # 其他示例
    'python trip_generator.py 2025 1 5000 "Amsterdam Centraal" --output real_distances.xlsx',
    'python trip_generator.py 2025 2 4000 "Rotterdam Centraal" --json',
]

def test_google_maps_connection(api_key: str = None):
    """测试Google Maps API连接"""
    # 如果没有提供API密钥，尝试从环境变量获取
    if not api_key:
        api_key = os.getenv('GOOGLE_MAPS_API_KEY')
        if api_key:
            print("✅ 从环境变量获取API密钥")
        else:
            print("❌ 未找到API密钥（环境变量或参数）")
            return False
    
    try:
        import googlemaps
        gmaps = googlemaps.Client(key=api_key)
        
        # 测试API调用
        result = gmaps.distance_matrix(
            origins=["Amsterdam, Netherlands"],
            destinations=["Utrecht, Netherlands"],
            mode="driving",
            units="metric"
        )
        
        if result['status'] == 'OK':
            distance_m = result['rows'][0]['elements'][0]['distance']['value']
            distance_km = distance_m / 1000
            print(f"✅ API连接成功！Amsterdam到Utrecht的距离: {distance_km:.1f}km")
            return True
        else:
            print(f"❌ API响应错误: {result['status']}")
            return False
            
    except ImportError:
        print("❌ 请先安装googlemaps包: pip install googlemaps")
        return False
    except Exception as e:
        print(f"❌ API连接失败: {e}")
        return False

def show_env_setup_instructions():
    """显示环境变量设置说明"""
    print("\n🔧 环境变量设置方法:")
    print("=" * 50)
    print("\n方法1: 临时设置（当前终端会话）")
    print('export GOOGLE_MAPS_API_KEY="your_actual_api_key"')
    
    print("\n方法2: 永久设置（添加到shell配置文件）")
    print("# 编辑 ~/.bashrc 或 ~/.zshrc")
    print('echo \'export GOOGLE_MAPS_API_KEY="your_actual_api_key"\' >> ~/.bashrc')
    print("source ~/.bashrc")
    
    print("\n方法3: 使用.env文件")
    print("1. 复制 .env.example 为 .env")
    print("2. 编辑 .env 文件，设置您的API密钥")
    print("3. 运行: source .env")
    
    print("\n方法4: Python脚本中设置")
    print('import os')
    print('os.environ["GOOGLE_MAPS_API_KEY"] = "your_actual_api_key"')

if __name__ == "__main__":
    print("Google Maps API使用说明")
    print("=" * 50)
    print("\n1. 获取API密钥:")
    print("   - 访问: https://console.cloud.google.com/")
    print("   - 创建项目并启用Distance Matrix API")
    print("   - 创建API密钥")
    
    show_env_setup_instructions()
    
    print("\n2. 使用示例:")
    for i, cmd in enumerate(example_commands, 1):
        if cmd:  # 跳过空行
            print(f"   {cmd}")
        else:
            print()
    
    print("\n3. 测试API连接:")
    
    # 首先检查环境变量
    env_api_key = os.getenv('GOOGLE_MAPS_API_KEY')
    if env_api_key:
        print(f"✅ 检测到环境变量中的API密钥: {env_api_key[:10]}...")
        test_choice = input("使用环境变量中的API密钥测试? (y/n): ").strip().lower()
        if test_choice == 'y':
            test_google_maps_connection()
        else:
            manual_key = input("请输入API密钥手动测试 (留空跳过): ").strip()
            if manual_key:
                test_google_maps_connection(manual_key)
            else:
                print("跳过API测试")
    else:
        print("ℹ️  未检测到环境变量 GOOGLE_MAPS_API_KEY")
        manual_key = input("请输入您的Google Maps API密钥测试 (留空跳过): ").strip()
        if manual_key:
            test_google_maps_connection(manual_key)
        else:
            print("跳过API测试") 