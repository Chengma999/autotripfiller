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
    
    # 比利时弗拉芒区域 (Vlaanderen)
    "Antwerpen", "Gent", "Brugge", "Leuven", "Mechelen", "Aalst", "Kortrijk",
    "Hasselt", "Sint-Niklaas", "Oostende", "Genk", "Roeselare", "Mouscron",
    "Verviers", "Turnhout", "Lokeren", "Beringen", "Sint-Truiden", "Brasschaat",
    "Schoten", "Deurne", "Wilrijk", "Edegem", "Kontich", "Aartselaar", "Hove",
    "Boechout", "Lint", "Niel", "Rumst", "Boom", "Schelle", "Hemiksem",
    "Hoboken", "Zwijndrecht", "Burcht", "Kruibeke", "Temse", "Bornem",
    "Puurs", "Sint-Amands", "Berlare", "Buggenhout", "Lebbeke", "Dendermonde",
    "Hamme", "Waasmunster", "Sint-Gillis-Waas", "Stekene", "Beveren",
    "Zele", "Lokeren", "Moerbeke", "Wachtebeke", "Zelzate", "Assenede",
    "Eeklo", "Kaprijke", "Sint-Laureins", "Waarschoot", "Knesselare",
    "Maldegem", "Aalter", "Beernem", "Oostkamp", "Bruges", "Damme",
    "Knokke-Heist", "Blankenberge", "De Haan", "Zuienkerke", "Jabbeke",
    "Oudenburg", "Gistel", "Ichtegem", "Torhout", "Koekelare", "Kortemark",
    "Hooglede", "Staden", "Moorslede", "Ledegem", "Menen", "Wervik",
    "Poperinge", "Vleteren", "Lo-Reninge", "Ieper", "Langemark-Poelkapelle",
    "Zonnebeke", "Geluveld", "Passendale", "Westrozebeke", "Staden",
    "Diksmuide", "Koekelare", "Kortemark", "Houthulst", "Merckem",
    "Klerken", "Woumen", "Vladslo", "Beerst", "Keiem", "Lampernisse",
    "Oostvleteren", "Westvleteren", "Elverdinge", "Brielen", "Dikkebus",
    "Voormezele", "Zillebeke", "Hollebeke", "Kemmel", "Wytschaete",
    "Messines", "Ploegsteert", "Comines", "Heuvelland", "Dranouter"
]

class TripGenerator:
    def __init__(self, year: int, quarter: int, target_km: int, start_location: str, google_api_key: str = None):
        self.year = year
        self.quarter = quarter
        self.target_km = target_km
        self.start_location = start_location
        
        # 优先使用传入的API密钥，其次使用环境变量
        self.google_api_key = google_api_key or os.getenv('GOOGLE_MAPS_API_KEY')
        self.trips = []
        
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
                print("⚠️  使用随机距离生成")
            except Exception as e:
                print(f"⚠️  Google Maps API连接失败: {e}")
                print("⚠️  使用随机距离生成")
        else:
            print("ℹ️  未提供Google Maps API密钥")
            print("   可以设置环境变量: export GOOGLE_MAPS_API_KEY='your_api_key'")
            print("   或使用 --google-api-key 参数")
        
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
        """计算距离（使用Google Maps API或随机值）"""
        if self.gmaps:
            try:
                # 使用Google Maps API计算实际距离
                result = self.gmaps.distance_matrix(
                    origins=[self.start_location],
                    destinations=[destination + ", Netherlands"],  # 添加国家信息提高准确性
                    mode="driving",
                    units="metric",
                    avoid="tolls"  # 避免收费道路
                )
                
                # 检查API响应
                if (result['status'] == 'OK' and 
                    result['rows'][0]['elements'][0]['status'] == 'OK'):
                    
                    # 提取距离（米）并转换为公里
                    distance_m = result['rows'][0]['elements'][0]['distance']['value']
                    distance_km = distance_m / 1000
                    return int(distance_km * 2)  # 来回距离
                else:
                    print(f"⚠️  无法获取到{destination}的距离信息，使用随机值")
                    return self._generate_random_distance()
                    
            except Exception as e:
                print(f"⚠️  Google Maps API调用失败: {e}")
                return self._generate_random_distance()
        else:
            return self._generate_random_distance()
    
    def _generate_random_distance(self) -> int:
        """生成随机距离"""
        # 荷兰境内的合理距离范围：50-400公里单程
        base_distance = random.randint(50, 400)
        return base_distance * 2  # 来回距离
    
    def generate_trips(self):
        """生成旅程记录"""
        start_date, end_date = self.get_quarter_dates()
        current_km = 0
        
        while current_km < self.target_km:
            # 随机选择目的地
            destination = random.choice(DUTCH_CITIES)
            
            # 生成随机日期
            trip_date = self.generate_random_date(start_date, end_date)
            
            # 计算距离
            distance = self.calculate_distance(destination)
            
            # 检查是否超过目标公里数
            if current_km + distance > self.target_km * 1.1:  # 允许10%的超出
                # 调整最后一次行程的距离
                distance = self.target_km - current_km
                if distance < 50:  # 如果剩余距离太少，跳过
                    break
            
            trip = {
                "date": trip_date.strftime("%d-%m-%Y"),
                "destination": destination,
                "description": "klant bezoeken",
                "total_distance": distance
            }
            
            self.trips.append(trip)
            current_km += distance
            
            # 如果达到或超过目标，结束
            if current_km >= self.target_km:
                break
        
        # 按日期排序
        self.trips.sort(key=lambda x: datetime.strptime(x["date"], "%d-%m-%Y"))
        
        return self.trips
    
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

def main():
    parser = argparse.ArgumentParser(
        description="Reisverslag Generator - Automatisch reisverslagen genereren en exporteren naar Excel",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Gebruiksvoorbeelden:
  python trip_generator.py 2025 1 5000 "de genestetlaan 299 den haag"
  python trip_generator.py 2025 1 5000 "de genestetlaan 299 den haag" --google-api-key YOUR_API_KEY
  python trip_generator.py 2024 4 3000 "Amsterdam Centraal" --output mijn_reisverslag.xlsx
        """
    )
    
    parser.add_argument('year', type=int, help='Jaar (bijvoorbeeld: 2025)')
    parser.add_argument('quarter', type=int, choices=[1, 2, 3, 4], 
                       help='Kwartaal (1, 2, 3, of 4)')
    parser.add_argument('target_km', type=int, help='Doel kilometers')
    parser.add_argument('start_location', type=str, help='Startlocatie')
    parser.add_argument('--output', '-o', type=str, help='Output Excel bestandsnaam')
    parser.add_argument('--json', action='store_true', help='Ook JSON bestand opslaan')
    parser.add_argument('--seed', type=int, help='Random seed (voor reproduceerbare resultaten)')
    parser.add_argument('--google-api-key', type=str, help='Google Maps API sleutel voor echte afstanden')
    
    args = parser.parse_args()
    
    # 设置随机种子（如果提供）
    if args.seed:
        random.seed(args.seed)
    
    print(f"🚗 Genereren van reisverslag voor {args.year} Q{args.quarter}...")
    print(f"📍 Startlocatie: {args.start_location}")
    print(f"🎯 Doel kilometers: {args.target_km}")
    if args.google_api_key:
        print(f"🗺️  Google Maps API: Ingeschakeld")
    else:
        print(f"🎲 Afstandberekening: Willekeurig")
    print("-" * 50)
    
    # 创建旅程生成器
    generator = TripGenerator(args.year, args.quarter, args.target_km, args.start_location, args.google_api_key)
    
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