import os
import random
import csv

# Create test directory if it doesn't exist
test_dir = "test_data/gtin14_test_files"
os.makedirs(test_dir, exist_ok=True)

# Function to generate a valid GTIN-14 code for a drug name
def generate_gtin14_for_drug(drug_name, drug_id):
    # Use consistent prefixes for specific drug types
    prefixes = {
        "注": "1491",  # Injection
        "点眼液": "1492",  # Eye drops
        "軟膏": "1493",  # Ointment
        "錠": "1494",  # Tablets
        "カプセル": "1495",  # Capsules
    }
    
    # Default prefix
    prefix = "1490"
    
    # Check drug name for specific types
    for key, val in prefixes.items():
        if key in drug_name:
            prefix = val
            break
    
    # Generate a consistent suffix based on drug_id
    # Use drug_id to ensure the same drug always gets the same code
    suffix = str(drug_id).zfill(4)
    
    # Add random digits to complete the 14-digit code
    remaining_digits = 14 - len(prefix) - len(suffix)
    middle = ''.join(str((ord(c) % 10)) for c in drug_name[:remaining_digits])
    
    # Pad with zeros if needed
    if len(middle) < remaining_digits:
        middle = middle + '0' * (remaining_digits - len(middle))
    elif len(middle) > remaining_digits:
        middle = middle[:remaining_digits]
    
    # Combine to form the GTIN-14 code
    return prefix + middle + suffix

# Read drug names from tmp_tana.csv
drug_data = []
try:
    with open("/home/ubuntu/attachments/820c8ac2-8819-43b1-a8fb-4b935f40c4c2/tmp_tana.csv", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader)  # Skip header
        for i, row in enumerate(reader):
            if len(row) >= 2 and row[1].strip():
                drug_id = int(row[0]) if row[0].strip() and row[0].isdigit() else i
                drug_data.append((drug_id, row[1]))
except Exception as e:
    print(f"Error reading tmp_tana.csv: {e}")
    # Fallback drug names if file can't be read
    drug_data = [
        (1, "ナノパスニードルⅡ（ＴＮ－３４０４）"),
        (3, "マイクロファインプラス３２Ｇ４ｍｍ"),
        (4, "アピドラ注ソロスター　３００単位"),
        (5, "インスリン　グラルギンＢＳ注カート｢リリー｣　３００単位"),
        (6, "インスリン　グラルギンＢＳ注ミリオペン｢リリー｣　３００単位"),
        (8, "インスリン　アスパルトＢＳ注ソロスターＮＲ｢サノフィ｣　３００単位"),
        (10, "トルリシティ皮下注０.７５ｍｇアテオス　０.５ｍＬ"),
        (11, "ヒューマログミックス５０注ミリオペン　３００単位"),
        (12, "ヒューマログ注ミリオペン　３００単位"),
        (13, "ランタスＸＲ注ソロスター　４５０単位"),
        (14, "ランタス注ソロスター　３００単位"),
        (15, "エイベリス点眼液０.００２％"),
        (16, "ジクロフェナクナトリウム坐剤２５ｍｇ｢日医工｣"),
        (18, "デュアック配合ゲル"),
        (19, "ベピオゲル２.５％"),
        (20, "ラタチモ配合点眼液｢ニットー｣"),
        (21, "新レシカルボン坐剤"),
        (22, "アンテベートクリーム０.０５％"),
        (23, "アンテベート軟膏０.０５％"),
        (25, "クロベタゾールプロピオン酸エステル軟膏０.０５％｢タイヨー｣"),
        (27, "サレックス軟膏０.０５％"),
        (29, "タプコム配合点眼液"),
        (30, "タプロス点眼液０.００１５％"),
        (31, "ディフェリンゲル０.１％"),
        (33, "デキサメタゾンプロピオン酸エステルローション０.１％｢ＭＹＫ｣"),
        (35, "トラマゾリン点鼻液０.１１８％｢ＡＦＰ｣"),
        (36, "ニュープロパッチ１８ｍｇ"),
        (37, "ニュープロパッチ４.５ｍｇ"),
        (38, "ネリゾナソリューション０.１％"),
        (41, "バクスミー点鼻粉末剤３ｍｇ"),
        (43, "テレミンソフト坐薬２ｍｇ"),
        (44, "ナウゼリン坐剤１０　１０ｍｇ"),
        (45, "ナウゼリン坐剤３０　３０ｍｇ"),
        (46, "ネリプロクト坐剤"),
        (47, "フルメタ軟膏　０.１％"),
        (48, "プロスタンディン軟膏０.００３％"),
        (49, "プロトピック軟膏０.０３％小児用"),
        (50, "ベタメタゾン酪酸エステルプロピオン酸エステルローション０.０５％｢ＭＹＫ｣"),
        (51, "マーデュオックス軟膏"),
        (52, "メサデルムクリーム０.１％")
    ]

# Generate GTIN-14 codes for each drug
gtin14_codes = []
for drug_id, drug_name in drug_data:
    gtin14_code = generate_gtin14_for_drug(drug_name, drug_id)
    gtin14_codes.append((gtin14_code, drug_name))

# Generate 10 test files
for i in range(1, 11):
    filename = f"{test_dir}/gtin14_test_{i}.csv"
    
    # Determine number of codes (between 2 and 30)
    num_codes = random.randint(2, 30)
    
    # Randomly select codes from the generated list
    selected_codes = random.sample(gtin14_codes, min(num_codes, len(gtin14_codes)))
    
    with open(filename, "w", encoding="utf-8") as f:
        # Write header
        f.write("GTIN\n")
        
        # Write selected GTIN-14 codes
        for code, _ in selected_codes:
            f.write(f"{code}\n")
    
    print(f"Created {filename} with {num_codes} GTIN-14 codes")

# Create a mapping file for reference
mapping_filename = f"{test_dir}/gtin14_drug_mapping.csv"
with open(mapping_filename, "w", encoding="utf-8") as f:
    f.write("GTIN,薬品名\n")
    for code, name in gtin14_codes:
        f.write(f"{code},{name}\n")

print(f"Created mapping file {mapping_filename}")
print("All test files generated successfully")
