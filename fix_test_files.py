import os

# Directory containing test files
test_dir = "test_data/gtin14_test_files"

# Process all test files
for i in range(1, 11):
    filename = f"{test_dir}/gtin14_test_{i}.csv"
    
    # Read file content
    with open(filename, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Filter out empty lines and ensure each line contains a valid 14-digit code
    valid_lines = []
    for line in lines:
        line = line.strip()
        if line and line.isdigit() and len(line) == 14:
            valid_lines.append(line)
    
    # Write back only valid lines
    with open(filename, 'w', encoding='utf-8') as f:
        for line in valid_lines:
            f.write(f"{line}\n")
    
    print(f"Fixed {filename}: {len(valid_lines)} valid GTIN-14 codes")

print("All test files fixed successfully")
