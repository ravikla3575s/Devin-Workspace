import os
import random
import csv

# Create test directory if it doesn't exist
test_dir = "test_data/gtin14_test_files"
os.makedirs(test_dir, exist_ok=True)

# Sample GTIN-14 prefixes
prefixes = ["1491", "1492", "1493", "1494", "1495"]

# Drug names from tmp_tana.csv
drug_names = [
    "ナノパスニードルⅡ（ＴＮ－３４０４）",
    "マイクロファインプラス３２Ｇ４ｍｍ",
    "アピドラ注ソロスター　３００単位",
    "インスリン　グラルギンＢＳ注カート｢リリー｣　３００単位",
    "インスリン　グラルギンＢＳ注ミリオペン｢リリー｣　３００単位",
    "インスリン　アスパルトＢＳ注ソロスターＮＲ｢サノフィ｣　３００単位",
    "トルリシティ皮下注０.７５ｍｇアテオス　０.５ｍＬ",
    "ヒューマログミックス５０注ミリオペン　３００単位",
    "ヒューマログ注ミリオペン　３００単位",
    "ランタスＸＲ注ソロスター　４５０単位",
    "ランタス注ソロスター　３００単位",
    "エイベリス点眼液０.００２％",
    "ジクロフェナクナトリウム坐剤２５ｍｇ｢日医工｣",
    "デュアック配合ゲル",
    "ベピオゲル２.５％",
    "ラタチモ配合点眼液｢ニットー｣",
    "新レシカルボン坐剤",
    "アンテベートクリーム０.０５％",
    "アンテベート軟膏０.０５％",
    "クロベタゾールプロピオン酸エステル軟膏０.０５％｢タイヨー｣",
    "サレックス軟膏０.０５％",
    "タプコム配合点眼液",
    "タプロス点眼液０.００１５％",
    "ディフェリンゲル０.１％",
    "デキサメタゾンプロピオン酸エステルローション０.１％｢ＭＹＫ｣",
    "トラマゾリン点鼻液０.１１８％｢ＡＦＰ｣",
    "ニュープロパッチ１８ｍｇ",
    "ニュープロパッチ４.５ｍｇ",
    "ネリゾナソリューション０.１％",
    "バクスミー点鼻粉末剤３ｍｇ",
    "テレミンソフト坐薬２ｍｇ",
    "ナウゼリン坐剤１０　１０ｍｇ",
    "ナウゼリン坐剤３０　３０ｍｇ",
    "ネリプロクト坐剤",
    "フルメタ軟膏　０.１％",
    "プロスタンディン軟膏０.００３％",
    "プロトピック軟膏０.０３％小児用",
    "ベタメタゾン酪酸エステルプロピオン酸エステルローション０.０５％｢ＭＹＫ｣",
    "マーデュオックス軟膏",
    "メサデルムクリーム０.１％"
]

# Generate 10 test files
for i in range(1, 11):
    filename = f"{test_dir}/gtin14_test_{i}.csv"
    
    # Determine number of codes (between 2 and 30)
    num_codes = random.randint(2, 30)
    
    with open(filename, "w", encoding="utf-8") as f:
        # Write header
        f.write("GTIN\n")
        
        # Generate random GTIN-14 codes
        for _ in range(num_codes):
            prefix = random.choice(prefixes)
            suffix = ''.join(random.choices('0123456789', k=10))
            gtin = prefix + suffix
            f.write(f"{gtin}\n")
    
    print(f"Created {filename} with {num_codes} GTIN-14 codes")

print("All test files generated successfully")
