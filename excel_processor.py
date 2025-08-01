import pandas as pd
import sys
from pathlib import Path

def process_excel_file(input_file, output_file=None):
    """
    Process an Excel file and generate a transformed output.
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file (optional)
    """
    try:
        # Read the Excel file
        df = pd.read_excel(input_file)
        
        print(f"Loaded Excel file: {input_file}")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        
        # Example transformations (modify as needed)
        # Remove rows with any missing values
        df_cleaned = df.dropna()
        
        # Add a sample transformation column
        if len(df_cleaned) > 0:
            df_cleaned = df_cleaned.copy()
            df_cleaned['processed_at'] = pd.Timestamp.now()
        
        # Generate output filename if not provided
        if output_file is None:
            input_path = Path(input_file)
            output_file = input_path.parent / f"{input_path.stem}_processed{input_path.suffix}"
        
        # Save the processed data
        df_cleaned.to_excel(output_file, index=False)
        
        print(f"Processed data saved to: {output_file}")
        print(f"Original rows: {len(df)}, Processed rows: {len(df_cleaned)}")
        
        return df_cleaned
        
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return None

def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_processor.py <input_file> [output_file]")
        print("Example: python excel_processor.py data.xlsx processed_data.xlsx")
        return
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    process_excel_file(input_file, output_file)

if __name__ == "__main__":
    main()