import os
import random

# Create test directory if it doesn't exist
test_dir = "test_data/gtin14_test_files"
os.makedirs(test_dir, exist_ok=True)

# Path to the large CSV file with medical product codes
csv_file = "/home/ubuntu/attachments/69aa9f71-efd1-40ff-b9ec-671e5aa696ce/.csv"

# Extract GTIN-14 codes from the CSV file
gtin14_codes = []
count = 0

print(f"Extracting GTIN-14 codes from {csv_file}...")
with open(csv_file, 'r', encoding='utf-8', errors='ignore') as f:
    for i, line in enumerate(f):
        if i == 0:  # Skip header
            continue
        
        cols = line.split(',')
        if len(cols) > 32:  # The column index for 販売包装単位コード is 32
            code = cols[32].strip()
            # Verify it's a 14-digit code
            if len(code) == 14 and code.isdigit() and code.startswith('1'):  # Most GTIN-14 codes start with 1
                gtin14_codes.append(code)
                count += 1
                if count % 1000 == 0:
                    print(f"Processed {count} codes...")

if not gtin14_codes:
    print("Error: No valid GTIN-14 codes found in the file.")
    exit(1)

print(f"Found {len(gtin14_codes)} valid GTIN-14 codes.")

# Generate 10 test files
for i in range(1, 11):
    filename = f"{test_dir}/gtin14_test_{i}.csv"
    
    # Determine number of codes (between 2 and 30)
    num_codes = random.randint(2, 30)
    
    # Randomly select codes from the list
    selected_codes = random.sample(gtin14_codes, min(num_codes, len(gtin14_codes)))
    
    with open(filename, "w", encoding='utf-8') as f:
        # No header as per user's request
        # Write selected GTIN-14 codes
        for code in selected_codes:
            f.write(f"{code}\n")
    
    print(f"Created {filename} with {len(selected_codes)} GTIN-14 codes")

print("All test files generated successfully")
