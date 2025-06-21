#!/usr/bin/env python3
"""
旅程记录生成器
自动生成旅程记录并导出到Excel文件
"""

import argparse
import json
import random
import pandas as pd
from datetime import datetime, timedelta
from typing import List, Dict
import os

# 自动加载.env文件
def load_env_file():
    """加载.env文件中的环境变量"""
    env_path = '.env'
    if os.path.exists(env_path):
        with open(env_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key.strip()] = value.strip()
        print("✅ .env文件已加载")
    else:
        print("⚠️  .env文件不存在")

# 在导入后立即加载环境变量
load_env_file()

# 扩展的荷兰城市、村庄和比利时弗拉芒区域列表
DUTCH_CITIES = [
    # 荷兰主要城市
    "Amsterdam", "Rotterdam", "Den Haag", "Utrecht", "Eindhoven", "Tilburg",
    "Groningen", "Almere", "Breda", "Nijmegen", "Enschede", "Haarlem",
    "Arnhem", "Zaanstad", "Amersfoort", "Apeldoorn", "Den Bosch", "Hoofddorp",
    "Maastricht", "Leiden", "Dordrecht", "Zoetermeer", "Zwolle", "Deventer",
    "Delft", "Alkmaar", "Leeuwarden", "Venlo", "Hilversum", "Heerlen",
    "Purmerend", "Roosendaal", "Schiedam", "Spijkenisse", "Alphen aan den Rijn",
    "Gouda", "Vlaardingen", "Zeist", "Katwijk", "Nieuwegein", "Lelystad",
    "Oosterhout", "Emmen", "Veenendaal", "Helmond", "De Bilt", "Capelle aan den IJssel",
    "Bergen op Zoom", "Roermond", "Oss", "Leidschendam", "Voorschoten", "Hoorn",
    "Vlissingen", "Ridderkerk", "Barendrecht", "Hendrik-Ido-Ambacht", "Papendrecht",
    "Sliedrecht", "Gorinchem", "Vianen", "Nieuwkoop", "Bodegraven", "Woerden",
    "Montfoort", "IJsselstein", "Kamerik", "Harmelen", "Ter Aar", "Alphen",
    "Boskoop", "Waddinxveen", "Zoeterwoude", "Leiderdorp",
    "Oegstgeest", "Voorhout", "Sassenheim", "Hillegom", "Lisse", "Teylingen",
    "Noordwijk", "Noordwijkerhout", "De Zilk", "Bennebroek", "Heemstede",
    "Zandvoort", "Bloemendaal", "Beverwijk", "Heemskerk", "Castricum", "Uitgeest",
    "Akersloot", "Limmen", "Heiloo", "Bergen", "Schagen", "Heerhugowaard",
    "Langedijk", "Graft-De Rijp", "Schermer", "Koggenland", "Drechterland",
    "Stede Broec", "Enkhuizen", "Medemblik", "Opmeer", "Hollands Kroon",
    
    # Duiven及其周边地区 (Gelderland东部)
    "Duiven", "Westervoort", "Zevenaar", "Didam", "Wehl", "Doesburg", "Doetinchem",
    "Angerlo", "Babberich", "Giesbeek", "Lathum", "Loo", "Groessen", "Pannerden",
    "Angeren", "Huissen", "Bemmel", "Elst", "Oosterhout", "Slijk-Ewijk", "Driel",
    "Heteren", "Valburg", "Zetten", "Hemmen", "Dodewaard", "Opheusden", "Kesteren",
    "Rhenen", "Wageningen", "Bennekom", "Ede", "Veenendaal", "Renswoude",
    "Woudenberg", "Scherpenzeel", "Barneveld", "Voorthuizen", "Kootwijkerbroek",
    "Garderen", "Kootwijk", "Radio Kootwijk", "Uddel", "Elspeet", "Nunspeet",
    "Harderwijk", "Hierden", "Putten", "Ermelo", "Horst", "Voorthuizen",
    
    # Achterhoek地区 (Duiven东南部)
    "Montferland", "Bergh", "Didam", "Wehl", "Doesburg", "Doetinchem", "Gaanderen",
    "Terborg", "Silvolde", "Ulft", "Gendringen", "Dinxperlo", "Aalten", "Bredevoort",
    "Winterswijk", "Woold", "Meddo", "Ratum", "Groenlo", "Lichtenvoorde", "Harreveld",
    "Eibergen", "Neede", "Borculo", "Ruurlo", "Vorden", "Warnsveld", "Lochem",
    "Gorssel", "Epse", "Deventer", "Bathmen", "Holten", "Rijssen", "Wierden",
    "Enter", "Delden", "Hengelo", "Enschede", "Oldenzaal", "Losser", "Denekamp",
    
    # Betuwe地区 (Duiven西南部)
    "Lingewaard", "Huissen", "Bemmel", "Gendt", "Angeren", "Doornenburg", "Haalderen",
    "Leuth", "Loo", "Pannerden", "Ressen", "Elst", "Oosterhout", "Slijk-Ewijk",
    "Driel", "Heteren", "Randwijk", "Herveld", "Valburg", "Zetten", "Hemmen",
    "Dodewaard", "Opheusden", "Kesteren", "IJzendoorn", "Ochten", "Echteld",
    "Lienden", "Maurik", "Buren", "Kerk-Avezaath", "Zoelen", "Ravenswaaij",
    "Tiel", "Kapel-Avezaath", "Wadenoijen", "Rumpt", "Geldermalsen", "Beesd",
    "Rhenoy", "Deil", "Enspijk", "Haaften", "Tuil", "Brakel", "Poederoijen",
    "Zaltbommel", "Kerkwijk", "Alphen", "Maasdriel", "Hedel", "Ammerzoden",
    "Rossum", "Hurwenen", "Alem", "Maren-Kessel", "Lith", "Oijen", "Teeffelen",
    
    # Veluwe地区 (Duiven北部)
    "Rheden", "Rozendaal", "Velp", "Dieren", "Laag-Soeren", "De Steeg", "Ellecom",
    "Spankeren", "Lieren", "Brummen", "Hall", "Eerbeek", "Loenen", "Beekbergen",
    "Vorchten", "Twello", "Wilp", "Teuge", "Ugchelen", "Hoenderloo", "Otterlo",
    "Ede", "Bennekom", "Wageningen", "Renkum", "Heelsum", "Doorwerth", "Oosterbeek",
    "Wolfheze", "Renkum", "Heveadorp", "Driel", "Randwijk", "Herveld-Onder",
    "Andelst", "Oosterhout", "Kesteren", "Opheusden", "Dodewaard", "Hemmen",
    
    # 荷兰村庄 (dorpen)
    "Volendam", "Marken", "Edam", "Monnickendam", "Broek in Waterland", "Oostzaan",
    "Wormer", "Jisp", "Neck", "Westzaan", "Krommenie", "Wormerveer", "Zaandijk",
    "Koog aan de Zaan", "Assendelft", "Oostknollendam", "Watergang", "Zuiderwoude",
    "Ransdorp", "Holysloot", "Zunderdorp", "Schellingwoude", "Durgerdam",
    "Muiden", "Muiderberg", "Weesp", "Diemen", "Ouder-Amstel", "Amstelveen",
    "Aalsmeer", "Kudelstaart", "Uithoorn", "De Kwakel", "Mijdrecht", "Wilnis",
    "Vinkeveen", "Waverveen", "Abcoude", "Baambrugge", "Loenen aan de Vecht",
    "Breukelen", "Kockengen", "Tienhoven", "Oud-Zuilen", "Zuilen", "Maarssen",
    "Maarssenbroek", "Nieuwersluis", "Loenersloot", "Portengen", "Westbroek",
    "Hollandsche Rading", "Lage Vuursche", "Den Dolder", "Huis ter Heide",
    "Driebergen-Rijsenburg", "Doorn", "Leersum", "Maarn", "Maarsbergen",
    "Woudenberg", "Scherpenzeel", "Renswoude", "Wijk bij Duurstede",
    "Langbroek", "Cothen", "Werkhoven", "Odijk", "Bunnik", "Houten",
    "Vianen", "Lexmond", "Hagestein", "Everdingen", "Zijderveld", "Schoonhoven",
    "Haastrecht", "Vlist", "Stolwijk", "Bergambacht", "Ammerstol", "Streefkerk",
    "Liesveld", "Nieuwpoort", "Langerak", "Ameide", "Tienhoven", "Lexmond",
    "Acquoy", "Asperen", "Heukelum", "Spijk", "Neerijnen", "Ophemert",
    "Varik", "Heesselt", "Rumpt", "Geldermalsen", "Beesd", "Rhenoy",
    "Deil", "Enspijk", "Haaften", "Tuil", "Brakel", "Poederoijen", "Zaltbommel",
    "Kerkwijk", "Alphen", "Maasdriel", "Hedel", "Ammerzoden", "Rossum",
    "Hurwenen", "Alem", "Maren-Kessel", "Lith", "Oijen", "Teeffelen",
    "Heesch", "Nistelrode", "Dinther", "Loosbroek", "Vorstenbosch",
    
    # 比利时弗拉芒区域 (Vlaanderen) - 添加 BE 标识
    "Antwerpen BE", "Gent BE", "Brugge BE", "Leuven BE", "Mechelen BE", "Aalst BE", "Kortrijk BE",
    "Hasselt BE", "Sint-Niklaas BE", "Oostende BE", "Genk BE", "Roeselare BE", "Mouscron BE",
    "Verviers BE", "Turnhout BE", "Lokeren BE", "Beringen BE", "Sint-Truiden BE", "Brasschaat BE",
    "Schoten BE", "Deurne BE", "Wilrijk BE", "Edegem BE", "Kontich BE", "Aartselaar BE", "Hove BE",
    "Boechout BE", "Lint BE", "Niel BE", "Rumst BE", "Boom BE", "Schelle BE", "Hemiksem BE",
    "Hoboken BE", "Zwijndrecht BE", "Burcht BE", "Kruibeke BE", "Temse BE", "Bornem BE",
    "Puurs BE", "Sint-Amands BE", "Berlare BE", "Buggenhout BE", "Lebbeke BE", "Dendermonde BE",
    "Hamme BE", "Waasmunster BE", "Sint-Gillis-Waas BE", "Stekene BE", "Beveren BE",
    "Zele BE", "Lokeren BE", "Moerbeke BE", "Wachtebeke BE", "Zelzate BE", "Assenede BE",
    "Eeklo BE", "Kaprijke BE", "Sint-Laureins BE", "Waarschoot BE", "Knesselare BE",
    "Maldegem BE", "Aalter BE", "Beernem BE", "Oostkamp BE", "Bruges BE", "Damme BE",
    "Knokke-Heist BE", "Blankenberge BE", "De Haan BE", "Zuienkerke BE", "Jabbeke BE",
    "Oudenburg BE", "Gistel BE", "Ichtegem BE", "Torhout BE", "Koekelare BE", "Kortemark BE",
    "Hooglede BE", "Staden BE", "Moorslede BE", "Ledegem BE", "Menen BE", "Wervik BE",
    "Poperinge BE", "Vleteren BE", "Lo-Reninge BE", "Ieper BE", "Langemark-Poelkapelle BE",
    "Zonnebeke BE", "Geluveld BE", "Passendale BE", "Westrozebeke BE", "Staden BE",
    "Diksmuide BE", "Koekelare BE", "Kortemark BE", "Houthulst BE", "Merckem BE",
    "Klerken BE", "Woumen BE", "Vladslo BE", "Beerst BE", "Keiem BE", "Lampernisse BE",
    "Oostvleteren BE", "Westvleteren BE", "Elverdinge BE", "Brielen BE", "Dikkebus BE",
    "Voormezele BE", "Zillebeke BE", "Hollebeke BE", "Kemmel BE", "Wytschaete BE",
    "Messines BE", "Ploegsteert BE", "Comines BE", "Heuvelland BE", "Dranouter BE"
]

class TripGenerator:
    def __init__(self, year: int, quarter: int, target_km: int, start_location: str, google_api_key: str = None):
        self.year = year
        self.quarter = quarter
        self.target_km = target_km
        self.start_location = start_location
        
        # 优先使用传入的API密钥，其次使用环境变量
        env_api_key = os.getenv('GOOGLE_MAPS_API_KEY')
        self.google_api_key = google_api_key or env_api_key
        self.trips = []
        
        # 调试信息
        print(f"🔍 API密钥检测:")
        print(f"   命令行参数: {'✅ 有' if google_api_key else '❌ 无'}")
        print(f"   环境变量: {'✅ 有' if env_api_key else '❌ 无'}")
        print(f"   最终使用: {'✅ 有' if self.google_api_key else '❌ 无'}")
        
        # 初始化Google Maps客户端（如果提供了API密钥）
        self.gmaps = None
        if self.google_api_key:
            try:
                import googlemaps
                self.gmaps = googlemaps.Client(key=self.google_api_key)
                print("✅ Google Maps API已连接")
                if google_api_key:
                    print("   (使用命令行参数)")
                else:
                    print("   (使用环境变量 GOOGLE_MAPS_API_KEY)")
            except ImportError:
                print("⚠️  警告: 请安装googlemaps包: pip install googlemaps")
                print("⚠️  将无法使用真实距离计算")
            except Exception as e:
                print(f"⚠️  Google Maps API连接失败: {e}")
                print("⚠️  将无法使用真实距离计算")
        else:
            print("❌ 未找到Google Maps API密钥!")
            print("   请使用以下方式之一:")
            print("   1. 设置环境变量: export GOOGLE_MAPS_API_KEY='your_api_key'")
            print("   2. 使用命令行参数: --google-api-key YOUR_API_KEY")
            print("   3. 确保.env文件存在且已加载")
            print("⚠️  将无法使用真实距离计算")
        
    def get_quarter_dates(self) -> List[datetime]:
        """获取指定季度的日期范围"""
        if self.quarter == 1:
            start_month, end_month = 1, 3
        elif self.quarter == 2:
            start_month, end_month = 4, 6
        elif self.quarter == 3:
            start_month, end_month = 7, 9
        else:  # quarter == 4
            start_month, end_month = 10, 12
            
        start_date = datetime(self.year, start_month, 1)
        
        # 获取季度结束日期
        if end_month == 12:
            end_date = datetime(self.year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = datetime(self.year, end_month + 1, 1) - timedelta(days=1)
            
        return start_date, end_date
    
    def generate_random_date(self, start_date: datetime, end_date: datetime) -> datetime:
        """在指定日期范围内生成随机日期"""
        time_between = end_date - start_date
        days_between = time_between.days
        random_days = random.randrange(days_between)
        return start_date + timedelta(days=random_days)
    
    def calculate_distance(self, destination: str) -> int:
        """计算距离（仅使用Google Maps API真实距离）"""
        if not self.gmaps:
            print(f"❌ 无Google Maps API密钥，无法计算到 {destination} 的距离")
            return None
        
        # 检查是否是比利时城市（带 BE 标识）
        is_belgian = destination.endswith(" BE")
        clean_destination = destination.replace(" BE", "") if is_belgian else destination
        
        # 根据城市类型设置地址格式
        if is_belgian:
            address_variants = [
                f"{clean_destination}, Belgium",
                f"{clean_destination}, België",
                f"{clean_destination}, BE",
                f"{clean_destination}"
            ]
        else:
            # 荷兰城市的地址格式
            address_variants = [
                f"{clean_destination}, Netherlands",
                f"{clean_destination}, Nederland", 
                f"{clean_destination}",
                f"{clean_destination}, Holland"
            ]
        
        for address in address_variants:
            try:
                print(f"🔍 尝试地址: {address}")
                result = self.gmaps.distance_matrix(
                    origins=[self.start_location],
                    destinations=[address],
                    mode="driving",
                    units="metric",
                    avoid="tolls"
                )
                
                # 检查API响应
                if (result['status'] == 'OK' and 
                    len(result['rows']) > 0 and
                    len(result['rows'][0]['elements']) > 0 and
                    result['rows'][0]['elements'][0]['status'] == 'OK'):
                    
                    # 提取距离信息
                    element = result['rows'][0]['elements'][0]
                    distance_m = element['distance']['value']
                    distance_text = element['distance']['text']
                    duration_text = element['duration']['text']
                    
                    one_way_km = distance_m / 1000
                    round_trip_km = int(one_way_km * 2)
                    
                    print(f"✅ 成功匹配地址: {address}")
                    print(f"📏 单程距离: {one_way_km:.1f}km ({distance_text})")
                    print(f"🔄 来回距离: {round_trip_km}km")
                    print(f"⏱️  行驶时间: {duration_text}")
                    
                    return round_trip_km
                else:
                    element_status = 'UNKNOWN'
                    if (len(result['rows']) > 0 and 
                        len(result['rows'][0]['elements']) > 0):
                        element_status = result['rows'][0]['elements'][0].get('status', 'UNKNOWN')
                    
                    print(f"⚠️  地址匹配失败: {address}")
                    print(f"    API状态: {result.get('status', 'UNKNOWN')}")
                    print(f"    元素状态: {element_status}")
                    continue
                    
            except Exception as e:
                print(f"⚠️  API调用异常: {address} - {e}")
                continue
        
        # 如果所有地址变体都失败了
        print(f"❌ 无法获取到 {destination} 的距离信息")
        print(f"    已尝试的地址格式: {address_variants}")
        return None
    
    def generate_trips(self):
        """生成旅程记录（仅使用真实距离，按距离分布）"""
        start_date, end_date = self.get_quarter_dates()
        current_km = 0
        failed_destinations = []
        destination_counts = {}  # 跟踪每个目的地的使用次数
        date_counts = {}  # 跟踪每个日期的行程次数
        
        # 距离分布目标
        target_short = int(self.target_km * 0.4)   # 40% < 150km (调整)
        target_medium = int(self.target_km * 0.4)  # 40% 150-300km  
        target_long = int(self.target_km * 0.2)    # 20% > 300km
        
        # 当前各类距离累计
        current_short = 0   # < 150km (调整)
        current_medium = 0  # 150-300km
        current_long = 0    # > 300km
        
        print(f"🎯 距离分布目标:")
        print(f"   短途 (<150km): {target_short}km (40%)")
        print(f"   中途 (150-300km): {target_medium}km (40%)")
        print(f"   长途 (>300km): {target_long}km (20%)")
        print("-" * 50)
        
        max_attempts = 1000  # 防止无限循环
        attempts = 0
        
        while current_km < self.target_km and attempts < max_attempts:
            attempts += 1
            
            # 确定当前需要的距离类型
            needed_type = self._determine_needed_distance_type(
                current_short, current_medium, current_long,
                target_short, target_medium, target_long
            )
            
            # 根据需要的距离类型选择合适的目的地
            destination = self._select_destination_by_distance_type(
                needed_type, failed_destinations, destination_counts
            )
            
            if destination is None:
                print("❌ 无法找到合适的目的地，停止生成")
                break
            
            # 生成随机日期，确保该日期的行程次数不超过2次
            trip_date = self._generate_valid_date(start_date, end_date, date_counts)
            
            if trip_date is None:
                print("⚠️  无法找到合适的日期（所有日期都已有2次行程），停止生成")
                break
            
            # 计算距离（仅使用真实距离）
            distance = self.calculate_distance(destination)
            
            if distance is None:
                failed_destinations.append(destination)
                print(f"⏭️  跳过目的地: {destination}")
                
                # 如果失败的目的地太多，停止生成
                if len(failed_destinations) > 50:
                    print("❌ 太多目的地无法获取距离，请检查网络连接或API密钥")
                    break
                continue
            
            # 更新目的地使用次数和日期使用次数
            destination_counts[destination] = destination_counts.get(destination, 0) + 1
            date_str = trip_date.strftime("%d-%m-%Y")
            date_counts[date_str] = date_counts.get(date_str, 0) + 1
            
            # 更新相应的距离累计
            if distance < 150:
                current_short += distance
                distance_type = "短途"
            elif distance <= 300:
                current_medium += distance
                distance_type = "中途"
            else:
                current_long += distance
                distance_type = "长途"
            
            trip = {
                "date": date_str,
                "destination": destination,
                "description": "klant bezoeken",
                "total_distance": distance
            }
            
            self.trips.append(trip)
            current_km += distance
            
            # 显示目的地使用次数和日期使用次数
            count_info = f"({destination_counts[destination]}/3)" if destination_counts[destination] > 1 else ""
            date_info = f"[{date_counts[date_str]}/2日程]" if date_counts[date_str] > 1 else ""
            print(f"✨ 添加行程: {destination} {count_info}({distance}km - {distance_type}) {date_info}")
            print(f"📊 累计距离: {current_km}km / 目标: {self.target_km}km")
            print(f"   短途: {current_short}km / {target_short}km")
            print(f"   中途: {current_medium}km / {target_medium}km") 
            print(f"   长途: {current_long}km / {target_long}km")
            print("-" * 50)
            
            # 达到目标距离时停止
            if current_km >= self.target_km:
                break
        
        if attempts >= max_attempts:
            print("⚠️  达到最大尝试次数，可能由于日期或目的地限制无法继续生成")
        
        if failed_destinations:
            print(f"\n⚠️  以下目的地无法获取距离: {failed_destinations[:10]}...")
        
        # 显示目的地使用统计
        self._print_destination_usage(destination_counts)
        
        # 显示日期使用统计
        self._print_date_usage(date_counts)
        
        # 显示最终分布
        self._print_final_distribution(current_short, current_medium, current_long)
        
        # 按日期排序
        self.trips.sort(key=lambda x: datetime.strptime(x["date"], "%d-%m-%Y"))
        
        return self.trips
    
    def _determine_needed_distance_type(self, current_short, current_medium, current_long,
                                      target_short, target_medium, target_long):
        """确定当前最需要的距离类型"""
        # 计算各类型的完成百分比
        short_ratio = current_short / target_short if target_short > 0 else 1
        medium_ratio = current_medium / target_medium if target_medium > 0 else 1
        long_ratio = current_long / target_long if target_long > 0 else 1
        
        # 优先选择完成度最低的类型
        if short_ratio <= medium_ratio and short_ratio <= long_ratio:
            return "short"
        elif medium_ratio <= long_ratio:
            return "medium"
        else:
            return "long"
    
    def _select_destination_by_distance_type(self, distance_type, failed_destinations, destination_counts=None):
        """根据距离类型选择合适的目的地"""
        if destination_counts is None:
            destination_counts = {}
        
        # 根据起始地点调整城市分类
        # 如果起点包含Duiven，使用Duiven周边的分类
        if "Duiven" in self.start_location:
            if distance_type == "short":
                # 短途：Duiven周边城市 (<150km往返)
                preferred_cities = [
                    # Duiven直接周边 (非常近)
                    "Arnhem", "Nijmegen", "Zevenaar", "Westervoort", "Doesburg", "Doetinchem",
                    "Huissen", "Bemmel", "Elst", "Wageningen", "Ede", "Rheden", "Velp",
                    "Dieren", "Brummen", "Zutphen", "Deventer", "Apeldoorn",
                    # Gelderland省内较近城市
                    "Zwolle", "Amersfoort", "Utrecht", "Nieuwegein", "Veenendaal",
                    "Barneveld", "Harderwijk", "Ermelo", "Putten", "Nunspeet",
                    # 扩展短途范围 (100-150km)
                    "Lelystad", "Almere", "Hilversum", "Gouda", "Woerden", "Montfoort"
                ]
            elif distance_type == "medium":
                # 中途：荷兰中部和西部城市 (150-300km往返)
                preferred_cities = [
                    "Amsterdam", "Rotterdam", "Den Haag", "Haarlem", "Leiden", "Delft",
                    "Alphen aan den Rijn", "Zoetermeer", "Dordrecht", "Vlaardingen",
                    "Schiedam", "Hoofddorp", "Alkmaar", "Hoorn", "Zaanstad", "Purmerend",
                    "Eindhoven", "Tilburg", "Breda", "Den Bosch", "Oss", "Helmond", 
                    "Roosendaal", "Bergen op Zoom", "Venlo", "Roermond"
                ]
            else:  # long
                # 长途：荷兰北部、南部远距离城市和比利时 (>300km往返)
                preferred_cities = [
                    "Groningen", "Leeuwarden", "Enschede", "Emmen", "Maastricht", "Heerlen",
                    "Sittard", "Geleen", "Kerkrade", "Brunssum",
                    "Antwerpen BE", "Gent BE", "Brugge BE", "Kortrijk BE", "Hasselt BE", 
                    "Leuven BE", "Mechelen BE", "Oostende BE", "Mouscron BE", "Sint-Niklaas BE", 
                    "Turnhout BE", "Genk BE", "Brasschaat BE", "Ledegem BE", "Ieper BE", 
                    "Poperinge BE", "Lo-Reninge BE", "Westrozebeke BE"
                ]
        else:
            # 默认分类（适用于海牙等西部城市）
            if distance_type == "short":
                # 短途：主要是海牙周边城市
                preferred_cities = [
                    "Delft", "Leidschendam", "Voorschoten", "Zoetermeer", "Rijswijk",
                    "Wassenaar", "Katwijk", "Noordwijk", "Leiden", "Alphen aan den Rijn",
                    "Gouda", "Bodegraven", "Woerden", "Vlaardingen", "Schiedam",
                    "Rotterdam", "Dordrecht", "Nieuwegein", "Utrecht", "Hoofddorp",
                    "Haarlem", "Amsterdam", "Hilversum"
                ]
            elif distance_type == "medium":
                # 中途：荷兰境内较远城市
                preferred_cities = [
                    "Eindhoven", "Tilburg", "Breda", "Bergen op Zoom", "Roosendaal",
                    "Den Bosch", "Oss", "Nijmegen", "Arnhem", "Apeldoorn", "Zwolle",
                    "Deventer", "Amersfoort", "Almere", "Lelystad", "Alkmaar",
                    "Hoorn", "Enkhuizen", "Medemblik", "Purmerend", "Zaanstad",
                    "Heerhugowaard", "Bergen", "Castricum", "Beverwijk"
                ]
            else:  # long
                # 长途：荷兰北部、东部和比利时城市（带 BE 标识）
                preferred_cities = [
                    "Groningen", "Leeuwarden", "Enschede", "Emmen", "Maastricht",
                    "Heerlen", "Venlo", "Roermond", "Helmod", "Antwerpen BE", "Gent BE",
                    "Brugge BE", "Kortrijk BE", "Hasselt BE", "Leuven BE", "Mechelen BE", "Oostende BE",
                    "Mouscron BE", "Sint-Niklaas BE", "Turnhout BE", "Genk BE", "Brasschaat BE",
                    "Ledegem BE", "Ieper BE", "Poperinge BE", "Lo-Reninge BE", "Westrozebeke BE"
                ]
        
        # 过滤出可用的城市（未失败且使用次数少于3次）
        available_cities = [
            city for city in preferred_cities 
            if (city not in failed_destinations and 
                destination_counts.get(city, 0) < 3)
        ]
        
        if not available_cities:
            # 如果首选城市都失败了或超过使用限制，从所有城市中选择
            available_cities = [
                city for city in DUTCH_CITIES 
                if (city not in failed_destinations and 
                    destination_counts.get(city, 0) < 3)
            ]
        
        if not available_cities:
            print("⚠️  所有目的地都已达到3次使用限制或无法访问")
            return None
        
        return random.choice(available_cities)
    
    def _generate_valid_date(self, start_date: datetime, end_date: datetime, date_counts: dict, max_attempts: int = 100) -> datetime:
        """生成一个有效的日期，确保该日期的行程次数不超过2次"""
        for _ in range(max_attempts):
            trip_date = self.generate_random_date(start_date, end_date)
            date_str = trip_date.strftime("%d-%m-%Y")
            
            # 检查该日期的行程次数是否少于2次
            if date_counts.get(date_str, 0) < 2:
                return trip_date
        
        # 如果随机生成失败，尝试找到第一个可用的日期
        current_date = start_date
        while current_date <= end_date:
            date_str = current_date.strftime("%d-%m-%Y")
            if date_counts.get(date_str, 0) < 2:
                return current_date
            current_date += timedelta(days=1)
        
        # 如果所有日期都已满，返回None
        return None
    
    def _print_final_distribution(self, current_short, current_medium, current_long):
        """打印最终的距离分布"""
        total = current_short + current_medium + current_long
        if total == 0:
            return
        
        short_percent = (current_short / total) * 100
        medium_percent = (current_medium / total) * 100
        long_percent = (current_long / total) * 100
        
        print(f"\n📊 最终距离分布:")
        print(f"   短途 (<150km): {current_short}km ({short_percent:.1f}%)")
        print(f"   中途 (150-300km): {current_medium}km ({medium_percent:.1f}%)")
        print(f"   长途 (>300km): {current_long}km ({long_percent:.1f}%)")
        print(f"   总计: {total}km")
    
    def export_to_excel(self, filename: str = None):
        """导出到Excel文件"""
        if not filename:
            filename = f"reisverslag_{self.year}_Q{self.quarter}.xlsx"
        
        # 创建DataFrame
        df = pd.DataFrame(self.trips)
        
        # 重命名列为荷兰语
        df = df.rename(columns={
            "date": "Datum",
            "destination": "Bestemming",
            "description": "Omschrijving",
            "total_distance": "Totale afstand (km)"
        })
        
        # 添加汇总信息
        total_distance = sum(trip["total_distance"] for trip in self.trips)
        summary_data = {
            "Datum": "Totaal",
            "Bestemming": f"{len(self.trips)} ritten",
            "Omschrijving": f"{self.year} Q{self.quarter}",
            "Totale afstand (km)": total_distance
        }
        
        # 添加汇总行
        df_summary = pd.DataFrame([summary_data])
        df = pd.concat([df, df_summary], ignore_index=True)
        
        # 导出到Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Reisverslag', index=False)
            
            # 格式化工作表
            worksheet = writer.sheets['Reisverslag']
            
            # 设置列宽
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Reisverslag geëxporteerd naar: {filename}")
        print(f"📊 Totaal: {len(self.trips)} ritten, {total_distance} km")
        
        return filename
    
    def save_json(self, filename: str = None):
        """保存为JSON文件（可选）"""
        if not filename:
            filename = f"reisverslag_{self.year}_Q{self.quarter}.json"
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.trips, f, ensure_ascii=False, indent=2)
        
        print(f"📄 JSON bestand opgeslagen als: {filename}")
        return filename

    def _print_destination_usage(self, destination_counts):
        """打印目的地使用统计"""
        if not destination_counts:
            return
        
        print(f"\n📍 目的地使用统计:")
        # 按使用次数排序
        sorted_destinations = sorted(destination_counts.items(), key=lambda x: x[1], reverse=True)
        
        for destination, count in sorted_destinations:
            if count > 1:
                print(f"   {destination}: {count}次")
        
        max_usage = max(destination_counts.values()) if destination_counts else 0
        total_destinations = len(destination_counts)
        print(f"   总共使用 {total_destinations} 个不同目的地")
        print(f"   最多重复次数: {max_usage}次")

    def _print_date_usage(self, date_counts):
        """打印日期使用统计"""
        if not date_counts:
            return
        
        print(f"\n📅 日期使用统计:")
        # 只显示有2次行程的日期
        dates_with_2_trips = {date: count for date, count in date_counts.items() if count == 2}
        
        if dates_with_2_trips:
            print(f"   有2次行程的日期: {len(dates_with_2_trips)}天")
            # 按日期排序显示前几个
            sorted_dates = sorted(dates_with_2_trips.items())
            for date, count in sorted_dates[:5]:  # 只显示前5个
                print(f"      {date}: {count}次")
            if len(sorted_dates) > 5:
                print(f"      ... 还有{len(sorted_dates) - 5}天")
        else:
            print(f"   所有日期都只有1次行程")
        
        total_days = len(date_counts)
        total_trips = sum(date_counts.values())
        avg_trips_per_day = total_trips / total_days if total_days > 0 else 0
        print(f"   总共使用 {total_days} 天")
        print(f"   平均每天行程数: {avg_trips_per_day:.1f}次")

def main():
    parser = argparse.ArgumentParser(
        description="Reisverslag Generator - Automatisch reisverslagen genereren en exporteren naar Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Gebruiksvoorbeelden:
  python trip_generator.py --year 2025 --quarter 1 --target-km 5000 --address "起始地址"
  python trip_generator.py --year 2025 --quarter 1 --target-km 5000 --address "起始地址" --google-api-key YOUR_API_KEY
  python trip_generator.py --year 2024 --quarter 4 --target-km 3000 --address "起始地址" --output mijn_reisverslag.xlsx
        """
    )
    
    parser.add_argument('--year', type=int, required=True, help='Jaar (bijvoorbeeld: 2025)')
    parser.add_argument('--quarter', type=int, choices=[1, 2, 3, 4], required=True,
                       help='Kwartaal (1, 2, 3, of 4)')
    parser.add_argument('--target-km', type=int, required=True, help='Doel kilometers')
    parser.add_argument('--address', type=str, required=True, help='Startlocatie')
    parser.add_argument('--output', '-o', type=str, help='Output Excel bestandsnaam')
    parser.add_argument('--json', action='store_true', help='Ook JSON bestand opslaan')
    parser.add_argument('--seed', type=int, help='Random seed (voor reproduceerbare resultaten)')
    parser.add_argument('--google-api-key', type=str, help='Google Maps API sleutel voor echte afstanden')
    
    args = parser.parse_args()
    
    # 设置随机种子（如果提供）
    if args.seed:
        random.seed(args.seed)
    
    print(f"🚗 Genereren van reisverslag voor {args.year} Q{args.quarter}...")
    print(f"📍 Startlocatie: {args.address}")
    print(f"🎯 Doel kilometers: {args.target_km}")
    if args.google_api_key:
        print(f"🗺️  Google Maps API: Ingeschakeld")
    else:
        print(f"🎲 Afstandberekening: Willekeurig")
    print("-" * 50)
    
    # 创建旅程生成器
    generator = TripGenerator(args.year, args.quarter, args.target_km, args.address, args.google_api_key)
    
    # 生成旅程
    trips = generator.generate_trips()
    
    print(f"✨ Succesvol {len(trips)} ritten gegenereerd")
    
    # 导出到Excel
    excel_file = generator.export_to_excel(args.output)
    
    # 可选：保存JSON文件
    if args.json:
        generator.save_json()
    
    print("\n📋 Reissamenvatting:")
    for i, trip in enumerate(trips[:5], 1):  # 显示前5次旅程
        print(f"  {i}. {trip['date']} - {trip['destination']} ({trip['total_distance']}km)")
    
    if len(trips) > 5:
        print(f"  ... en nog {len(trips) - 5} ritten")
    
    total_km = sum(trip['total_distance'] for trip in trips)
    print(f"\n🏁 Totale kilometers: {total_km}km (doel: {args.target_km}km)")

if __name__ == "__main__":
    main() 