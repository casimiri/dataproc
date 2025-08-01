import pandas as pd
import sys
import re
from pathlib import Path

def clean_dose_value(value):
    """Remove units from dose values (e.g., '100 Gy' -> '100')"""
    if pd.isna(value) or value == '':
        return value
    
    # Convert to string and remove common units
    value_str = str(value).strip()
    # Remove units like Gy, gr, Gray, etc. (case insensitive)
    cleaned = re.sub(r'\s*(Gy|gr|Gray|kGy|rad)\s*$', '', value_str, flags=re.IGNORECASE)
    
    try:
        # Try to convert to float if it's a number
        return float(cleaned) if cleaned else ''
    except ValueError:
        return cleaned

def process_variety_names(variety_string):
    """Split variety names by common separators and clean them"""
    if pd.isna(variety_string) or variety_string == '':
        return ['']
    
    # Split by common separators: comma, semicolon, pipe, 'and', '&'
    varieties = re.split(r'[,;|]|\sand\s|\s&\s', str(variety_string))
    
    # Clean up each variety name
    cleaned_varieties = []
    for variety in varieties:
        cleaned = variety.strip()
        if cleaned:
            cleaned_varieties.append(cleaned)
    
    return cleaned_varieties if cleaned_varieties else ['']

def parse_address_field(address_str):
    """Parse address field into comprehensive components including names, contact info, organization"""
    result = {
        'FirstName': '', 'LastName': '', 'Phone': '', 'Email': '',
        'Name of organization': '', 'Type of organization': '',
        'Street': '', 'POBox': '', 'City': '', 'Country': ''
    }
    
    if pd.isna(address_str) or address_str == '':
        return result
    
    address = str(address_str).strip()
    
    # Extract email addresses
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    emails = re.findall(email_pattern, address)
    if emails:
        result['Email'] = emails[0]
        address = re.sub(email_pattern, '', address)
    
    # Extract phone numbers (various formats)
    phone_patterns = [
        r'\+?\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{1,4}[-.\s]?\d{1,9}',
        r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',
        r'\d{10,}'
    ]
    for pattern in phone_patterns:
        phones = re.findall(pattern, address)
        if phones:
            result['Phone'] = phones[0].strip()
            address = re.sub(pattern, '', address, count=1)
            break
    
    # Organization keywords for classification
    org_keywords = {
        'university': 'Academic',
        'institute': 'Research',
        'research': 'Research',
        'laboratory': 'Research',
        'lab': 'Research',
        'college': 'Academic',
        'school': 'Academic',
        'department': 'Government',
        'ministry': 'Government',
        'company': 'Commercial',
        'corp': 'Commercial',
        'ltd': 'Commercial',
        'inc': 'Commercial',
        'foundation': 'Non-profit',
        'center': 'Research',
        'centre': 'Research'
    }
    
    # Split by common separators
    parts = re.split(r'[,\n\r]+', address)
    parts = [p.strip() for p in parts if p.strip()]
    
    # Extract organization name and type
    for part in parts:
        part_lower = part.lower()
        for keyword, org_type in org_keywords.items():
            if keyword in part_lower:
                result['Name of organization'] = part
                result['Type of organization'] = org_type
                break
        if result['Name of organization']:
            break
    
    # Extract names (assume first part might contain person name)
    if parts:
        first_part = parts[0]
        # Look for patterns like "Dr. John Smith" or "Smith, John"
        name_words = first_part.split()
        if len(name_words) >= 2:
            # Skip titles
            titles = ['dr', 'prof', 'mr', 'mrs', 'ms', 'professor', 'doctor']
            filtered_words = [w for w in name_words if w.lower().strip('.') not in titles]
            
            if len(filtered_words) >= 2:
                result['FirstName'] = filtered_words[0]
                result['LastName'] = filtered_words[1]
            elif len(filtered_words) == 1:
                result['LastName'] = filtered_words[0]
    
    # Extract address components
    address_parts = []
    for part in parts:
        if not result['Name of organization'] or part != result['Name of organization']:
            # Skip parts that look like names or contact info
            if not (result['FirstName'] and result['FirstName'] in part):
                address_parts.append(part)
    
    # Parse P.O. Box
    for part in address_parts:
        if re.search(r'p\.?o\.?\s*box|post\s*office\s*box', part.lower()):
            result['POBox'] = part
            address_parts.remove(part)
            break
    
    # Assign remaining parts to Street, City, Country
    if address_parts:
        if len(address_parts) >= 3:
            result['Street'] = address_parts[0]
            result['City'] = address_parts[1]
            result['Country'] = address_parts[-1]
        elif len(address_parts) == 2:
            result['City'] = address_parts[0]
            result['Country'] = address_parts[1]
        elif len(address_parts) == 1:
            # Try to determine if it's a country or city
            common_countries = ['usa', 'canada', 'uk', 'australia', 'germany', 'france', 'japan', 'china', 'india', 'brazil']
            if any(country in address_parts[0].lower() for country in common_countries):
                result['Country'] = address_parts[0]
            else:
                result['City'] = address_parts[0]
    
    return result

def extract_treatment_type(dose_str):
    """Extract treatment type from dose field (e.g., GAMMA, ELECTRON, etc.)"""
    if pd.isna(dose_str) or dose_str == '':
        return ''
    
    dose_text = str(dose_str).upper()
    
    # Common treatment types
    treatment_keywords = [
        'GAMMA', 'ELECTRON', 'X-RAY', 'NEUTRON', 'PROTON', 'BETA', 'ALPHA',
        'ION', 'BEAM', 'RADIATION', 'IRRADIATION', 'EMS', 'CHEMICAL'
    ]
    
    for treatment in treatment_keywords:
        if treatment in dose_text:
            return treatment
    
    # Default to GAMMA if numbers are present (common for radiation)
    if re.search(r'\d', dose_text):
        return 'GAMMA'
    
    return ''

def process_dose_field(dose_str):
    """Process dose field and split into multiple dose columns"""
    doses = {'dose 1': '', 'dose 2': '', 'dose 3': '', 'dose 4': '', 'dose 5': '',
             'dose 6': '', 'dose 7': '', 'dose 8': '', 'dose 9': '', 'dose 10': ''}
    
    if pd.isna(dose_str) or dose_str == '':
        return doses
    
    # Split by common separators and clean
    dose_values = re.split(r'[,;|]|\sand\s|\s&\s', str(dose_str))
    
    for i, dose in enumerate(dose_values[:10]):  # Max 10 doses
        cleaned_dose = clean_dose_value(dose.strip())
        if cleaned_dose != '':
            doses[f'dose {i+1}'] = cleaned_dose
    
    return doses

def get_latin_name(plant_name):
    """Get Latin name based on common plant knowledge"""
    if pd.isna(plant_name) or plant_name == '':
        return ''
    
    plant_lower = str(plant_name).lower()
    
    # Common plant name to Latin name mapping
    latin_mapping = {
        'rice': 'Oryza sativa',
        'wheat': 'Triticum aestivum',
        'maize': 'Zea mays',
        'corn': 'Zea mays',
        'barley': 'Hordeum vulgare',
        'soybean': 'Glycine max',
        'soya': 'Glycine max',
        'tomato': 'Solanum lycopersicum',
        'potato': 'Solanum tuberosum',
        'cotton': 'Gossypium hirsutum',
        'sunflower': 'Helianthus annuus',
        'bean': 'Phaseolus vulgaris',
        'pea': 'Pisum sativum',
        'chickpea': 'Cicer arietinum',
        'lentil': 'Lens culinaris',
        'sesame': 'Sesamum indicum',
        'millet': 'Pennisetum glaucum',
        'sorghum': 'Sorghum bicolor',
        'oat': 'Avena sativa',
        'rye': 'Secale cereale',
        'cassava': 'Manihot esculenta',
        'sweet potato': 'Ipomoea batatas',
        'yam': 'Dioscorea spp.',
        'banana': 'Musa spp.',
        'apple': 'Malus domestica',
        'orange': 'Citrus sinensis',
        'lemon': 'Citrus limon',
        'mango': 'Mangifera indica',
        'coconut': 'Cocos nucifera',
        'palm': 'Elaeis guineensis',
        'sugarcane': 'Saccharum officinarum',
        'tobacco': 'Nicotiana tabacum',
        'coffee': 'Coffea arabica',
        'tea': 'Camellia sinensis',
        'pepper': 'Capsicum annuum',
        'chili': 'Capsicum annuum',
        'onion': 'Allium cepa',
        'garlic': 'Allium sativum',
        'carrot': 'Daucus carota',
        'cabbage': 'Brassica oleracea',
        'lettuce': 'Lactuca sativa',
        'spinach': 'Spinacia oleracea'
    }
    
    # Check for exact matches first
    for common_name, latin_name in latin_mapping.items():
        if common_name in plant_lower:
            return latin_name
    
    return ''

def classify_species_type(plant_name, material_name=''):
    """Classify the type of species (Seed, Cutting, etc.)"""
    if pd.isna(plant_name) and pd.isna(material_name):
        return 'Seed'  # Default
    
    plant_text = str(plant_name).lower() if not pd.isna(plant_name) else ''
    material_text = str(material_name).lower() if not pd.isna(material_name) else ''
    combined_text = plant_text + ' ' + material_text
    
    # Classification keywords
    if any(word in combined_text for word in ['seed', 'grain', 'kernel']):
        return 'Seed'
    elif any(word in combined_text for word in ['cutting', 'stem', 'branch']):
        return 'Cutting'
    elif any(word in combined_text for word in ['leaf', 'leaves']):
        return 'Leaf'
    elif any(word in combined_text for word in ['root', 'tuber', 'bulb']):
        return 'Root/Tuber'
    elif any(word in combined_text for word in ['fruit', 'berry']):
        return 'Fruit'
    elif any(word in combined_text for word in ['pollen']):
        return 'Pollen'
    elif any(word in combined_text for word in ['tissue', 'callus']):
        return 'Tissue Culture'
    else:
        return 'Seed'  # Default assumption

def process_excel_file(input_file, output_file=None):
    """
    Process an Excel file according to specific dataset requirements.
    
    Input columns: Date Received, Entry No, Material, Plant Name, Dose, Date sent back, Address
    
    Rules:
    1. Create one row per variety name (from Material field)
    2. Duplicate all other values for each variety
    3. Remove duplicate records with same variety name
    4. Remove units from dose columns
    5. Parse address into separate components
    
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
        
        # No longer need simple column mapping - using intelligent parsing instead
        
        # Define output columns in exact order
        output_columns = [
            'DateReceived', 'IDAssigned', 'FirstName', 'LastName', 'Phone', 'Email',
            'Name of organization', 'Type of organization', 'Street', 'POBox', 'City', 'Country',
            'Treatment', 'choose Shrinkwrap', 'Total number of bags', 'Target trait(s)',
            'cooperation with the Joint FAO/IAEA', 'which type of project', 'Type species',
            'Common Name species', 'Latin Name species', 'Variety Name species', 'Samples quantity species',
            'dose 1', 'dose 2', 'dose 3', 'dose 4', 'dose 5', 'dose 6', 'dose 7', 'dose 8', 'dose 9', 'dose 10'
        ]
        
        # Process each row
        processed_rows = []
        
        for _, row in df.iterrows():
            # Get varieties from Material column and split them
            material_col = None
            for col in df.columns:
                if 'material' in col.lower():
                    material_col = col
                    break
            
            if material_col:
                varieties = process_variety_names(row[material_col])
            else:
                varieties = ['']
            
            # Create one row per variety
            for variety in varieties:
                new_row = {}
                
                # Basic field mappings - handle flexible column name matching
                date_received_col = None
                entry_no_col = None
                
                # Find Date Received column (flexible matching)
                for col in df.columns:
                    if 'date' in col.lower() and 'received' in col.lower():
                        date_received_col = col
                        break
                
                # Find Entry No column (flexible matching)
                for col in df.columns:
                    if 'entry' in col.lower() and 'no' in col.lower():
                        entry_no_col = col
                        break
                    elif col.lower().strip() in ['entry no', 'entryno', 'entry_no', 'id', 'entry id']:
                        entry_no_col = col
                        break
                
                if date_received_col:
                    new_row['DateReceived'] = row[date_received_col]
                if entry_no_col:
                    new_row['IDAssigned'] = row[entry_no_col]
                
                # Address parsing for names, contact info, organization, location
                address_col = None
                for col in df.columns:
                    if 'address' in col.lower():
                        address_col = col
                        break
                
                if address_col:
                    address_data = parse_address_field(row[address_col])
                    new_row.update(address_data)
                
                # Plant and variety information
                plant_name_col = None
                for col in df.columns:
                    if 'plant' in col.lower() and 'name' in col.lower():
                        plant_name_col = col
                        break
                
                if plant_name_col:
                    plant_name = row[plant_name_col]
                    new_row['Common Name species'] = plant_name
                    new_row['Latin Name species'] = get_latin_name(plant_name)
                    new_row['Type species'] = classify_species_type(plant_name, variety)
                
                new_row['Variety Name species'] = variety
                
                # Dose and treatment information
                dose_col = None
                for col in df.columns:
                    if 'dose' in col.lower():
                        dose_col = col
                        break
                
                if dose_col:
                    dose_data = process_dose_field(row[dose_col])
                    new_row.update(dose_data)
                    new_row['Treatment'] = extract_treatment_type(row[dose_col])
                
                processed_rows.append(new_row)
        
        # Create new dataframe from processed rows
        df_processed = pd.DataFrame(processed_rows)
        
        # Remove duplicates based on key fields
        if 'IDAssigned' in df_processed.columns and 'Variety Name species' in df_processed.columns:
            df_processed = df_processed.drop_duplicates(subset=['IDAssigned', 'Variety Name species'], keep='first')
        else:
            df_processed = df_processed.drop_duplicates()
        
        # Create final dataframe with all required columns
        final_df = pd.DataFrame()
        
        for output_col in output_columns:
            if output_col in df_processed.columns:
                final_df[output_col] = df_processed[output_col]
            else:
                # Create empty column if not found
                final_df[output_col] = ''
        
        # Generate output filename if not provided
        if output_file is None:
            input_path = Path(input_file)
            output_file = input_path.parent / f"{input_path.stem}_processed{input_path.suffix}"
        
        # Save the processed data with exact column headers
        final_df.to_excel(output_file, index=False)
        
        print(f"Processed data saved to: {output_file}")
        print(f"Original rows: {len(df)}, Processed rows: {len(final_df)}")
        print(f"Material column found: {material_col}")
        print(f"Output columns: {list(final_df.columns)}")
        print("Intelligent parsing applied for names, addresses, treatments, and species classification")
        
        return final_df
        
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