import pandas as pd
import os

# Create a dummy Excel file
data = {
    'Model': ['A1', 'B2', 'C3'],
    'Qty': [10, 20, 30],
    'Price': [100, 200, 300]
}
df = pd.DataFrame(data)
excel_path = "test_input.xlsx"
df.to_excel(excel_path, index=False)

print(f"Created {excel_path}")

# Simulate app logic
try:
    # 1. Get headers
    input_df = pd.read_excel(excel_path)
    headers = [str(c).strip() for c in input_df.columns]
    print(f"Headers: {headers}")
    
    # 2. Extract data (simulate user selecting all columns)
    selected_headers = headers
    extracted_data = []
    for _, row in input_df.iterrows():
        row_data = {}
        for h in selected_headers:
            if h in input_df.columns:
                val = row[h]
                if pd.notna(val):
                    row_data[h] = val
        if row_data:
            extracted_data.append(row_data)
            
    print(f"Extracted {len(extracted_data)} items.")
    print("Sample item:", extracted_data[0])
    
    # Cleanup
    if os.path.exists(excel_path):
        os.unlink(excel_path)
        print("Cleaned up test file.")
        
except Exception as e:
    print(f"Verification failed: {e}")
