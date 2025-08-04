import pandas as pd
import sys
import re
import os
import time
import json
from pathlib import Path
from openai import OpenAI
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def format_date_received(date_value):
    """Format date to YYYY.MM.DD format"""
    if pd.isna(date_value) or date_value == '':
        return ''
    
    try:
        # Handle various date formats
        if isinstance(date_value, pd.Timestamp):
            return date_value.strftime('%Y.%m.%d')
        elif isinstance(date_value, str):
            # Try to parse string dates
            date_obj = pd.to_datetime(date_value, errors='coerce')
            if not pd.isna(date_obj):
                return date_obj.strftime('%Y.%m.%d')
        else:
            # Try to convert other types to datetime
            date_obj = pd.to_datetime(date_value, errors='coerce')
            if not pd.isna(date_obj):
                return date_obj.strftime('%Y.%m.%d')
    except:
        pass
    
    return str(date_value)  # Return original if formatting fails

def clean_dose_value(value):
    """Extract only numeric values from dose fields (e.g., '100 Gy' -> '100')"""
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and extract all numbers (including decimals)
    value_str = str(value).strip()
    
    # Find all numeric values (including decimals)
    numbers = re.findall(r'\d+\.?\d*', value_str)
    
    if numbers:
        # Return the first number found
        try:
            return float(numbers[0]) if '.' in numbers[0] else int(numbers[0])
        except ValueError:
            return ''
    
    return ''

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
    
    # Extract names first (we need them to filter from organization name later)
    for part in parts:
        first_part = part
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
        break  # Only process first part for names
    
    # Extract organization name and type
    for part in parts:
        part_lower = part.lower()
        for keyword, org_type in org_keywords.items():
            if keyword in part_lower:
                # Remove titles, country names, and person names from organization name
                org_name = part
                titles = ['ph.d', 'phd', 'dr', 'dr.', 'mr', 'mr.', 'mrs', 'mrs.', 'ms', 'ms.', 'prof', 'prof.', 'professor', 'doctor']
                countries = ['usa', 'united states', 'us', 'canada', 'uk', 'united kingdom', 'australia', 'germany', 'france', 'japan', 'china', 'india', 'brazil', 'russia', 'italy', 'spain', 'mexico', 'argentina', 'south africa', 'egypt', 'nigeria', 'kenya', 'ghana', 'morocco', 'algeria', 'tunisia', 'ethiopia', 'sudan', 'tanzania', 'uganda', 'zambia', 'zimbabwe', 'botswana', 'namibia', 'angola', 'mozambique', 'madagascar', 'mauritius', 'seychelles', 'thailand', 'vietnam', 'philippines', 'indonesia', 'malaysia', 'singapore', 'south korea', 'north korea', 'taiwan', 'hong kong', 'myanmar', 'cambodia', 'laos', 'bangladesh', 'pakistan', 'afghanistan', 'iran', 'iraq', 'turkey', 'israel', 'palestine', 'jordan', 'lebanon', 'syria', 'saudi arabia', 'uae', 'qatar', 'kuwait', 'oman', 'yemen', 'bahrain', 'nepal', 'bhutan', 'sri lanka', 'maldives', 'poland', 'czech republic', 'slovakia', 'hungary', 'romania', 'bulgaria', 'croatia', 'serbia', 'bosnia', 'montenegro', 'albania', 'greece', 'cyprus', 'malta', 'portugal', 'netherlands', 'belgium', 'luxembourg', 'switzerland', 'austria', 'denmark', 'sweden', 'norway', 'finland', 'iceland', 'ireland', 'estonia', 'latvia', 'lithuania', 'belarus', 'ukraine', 'moldova', 'georgia', 'armenia', 'azerbaijan', 'kazakhstan', 'uzbekistan', 'turkmenistan', 'kyrgyzstan', 'tajikistan', 'mongolia', 'chile', 'peru', 'ecuador', 'colombia', 'venezuela', 'bolivia', 'paraguay', 'uruguay', 'guyana', 'suriname', 'cuba', 'jamaica', 'haiti', 'dominican republic', 'puerto rico', 'trinidad', 'barbados', 'bahamas', 'belize', 'costa rica', 'panama', 'guatemala', 'honduras', 'el salvador', 'nicaragua', 'new zealand', 'fiji', 'papua new guinea', 'solomon islands', 'vanuatu', 'samoa', 'tonga', 'palau', 'micronesia', 'marshall islands', 'kiribati', 'tuvalu', 'nauru']
                
                # Get first/last names to exclude
                first_name = result.get('FirstName', '').lower()
                last_name = result.get('LastName', '').lower()
                
                words = org_name.split()
                filtered_words = []
                for word in words:
                    word_clean = word.lower().strip('.,')
                    if (word_clean not in titles and 
                        word_clean not in countries and
                        word_clean != first_name and 
                        word_clean != last_name):
                        # Also check if it's part of a multi-word country name
                        skip_word = False
                        for country in countries:
                            if len(country.split()) > 1 and word_clean in country.split():
                                skip_word = True
                                break
                        if not skip_word:
                            filtered_words.append(word)
                result['Name of organization'] = ' '.join(filtered_words).strip()
                result['Type of organization'] = org_type
                break
        if result['Name of organization']:
            break
    
    
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

# OpenAI client initialization
client = None
if os.getenv('OPENAI_API_KEY'):
    client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

def call_openai_with_retry(prompt, max_tokens=300, max_retries=3, delay=1):
    """Call OpenAI API with retry logic and rate limiting"""
    if not client:
        print("Warning: OpenAI API key not found. Using fallback methods.")
        return None
    
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are an expert data analyst specializing in botanical and geographical information extraction. Analyze all provided data to extract accurate field values. Always respond with valid JSON only."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=max_tokens,
                temperature=0.1
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            print(f"OpenAI API call failed (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                time.sleep(delay * (2 ** attempt))  # Exponential backoff
            else:
                print("All OpenAI API attempts failed. Using fallback methods.")
                return None
    return None

def extract_all_fields_openai(row_data):
    """Extract all output fields from complete row data using OpenAI API"""
    if not row_data:
        return {}
    
    # Convert row data to a readable format
    row_text = "\n".join([f"{col}: {val}" for col, val in row_data.items() if not pd.isna(val) and val != ''])
    
    prompt = f"""Analyze this Excel row data and extract the following fields. Consider ALL the provided information to determine each field value accurately.

Row Data:
{row_text}

Extract these fields and return as JSON:
{{
  "FirstName": "person's first name",
  "LastName": "person's last name", 
  "Phone": "phone number",
  "Email": "email address",
  "Name_of_organization": "organization/institution name",
  "Type_of_organization": "Academic/Research/Government/Commercial/Non-profit",
  "Street": "street address",
  "POBox": "P.O. Box or BP if present",
  "City": "city name",
  "Country": "country name (standardized)",
  "Treatment": "treatment type (GAMMA/ELECTRON/X-RAY/etc)",
  "Common_Name_species": "common/vernacular name of plant",
  "Latin_Name_species": "scientific/Latin name (Genus species)",
  "Variety_Name_species": "variety/cultivar name",
  "Type_species": "Seed/Cutting/Leaf/Root/Fruit/etc"
}}

Use empty strings for fields that cannot be determined. Be precise and use standard naming conventions."""
    
    response = call_openai_with_retry(prompt, max_tokens=500)
    if response:
        try:
            # Clean response to ensure valid JSON
            cleaned_response = response.strip()
            if cleaned_response.startswith('```json'):
                cleaned_response = cleaned_response[7:-3]
            elif cleaned_response.startswith('```'):
                cleaned_response = cleaned_response[3:-3]
            
            # Remove trailing commas that cause JSON parsing errors
            import re
            cleaned_response = re.sub(r',(\s*[}\\]])', r'\\1', cleaned_response)
            # Also handle trailing comma at end of object
            cleaned_response = re.sub(r',(\s*})', r'\\1', cleaned_response)
            
            data = json.loads(cleaned_response)
            
            # Standardize field names to match output format
            standardized = {
                'FirstName': data.get('FirstName', ''),
                'LastName': data.get('LastName', ''),
                'Phone': data.get('Phone', ''),
                'Email': data.get('Email', ''),
                'Name of organization': data.get('Name_of_organization', ''),
                'Type of organization': data.get('Type_of_organization', ''),
                'Street': data.get('Street', ''),
                'POBox': data.get('POBox', ''),
                'City': data.get('City', ''),
                'Country': data.get('Country', ''),
                'Treatment': data.get('Treatment', ''),
                'Common Name species': data.get('Common_Name_species', ''),
                'Latin Name species': data.get('Latin_Name_species', ''),
                'Variety Name species': data.get('Variety_Name_species', ''),
                'Type species': data.get('Type_species', '')
            }
            
            return standardized
        except json.JSONDecodeError as e:
            print(f"JSON decode error: {e}")
            print(f"Response was: {response}")
    
    return {}

def get_plant_info_openai(plant_name, variety_name=''):
    """Get plant information including Latin name and standardized common name using OpenAI API"""
    if pd.isna(plant_name) or plant_name == '':
        return {'latin_name': '', 'common_name': '', 'variety_name': variety_name}
    
    plant_text = f"{plant_name} {variety_name}".strip()
    
    prompt = f"""Identify this plant and provide botanical information. Return only a JSON object with these keys:
- 'latin_name': The scientific/Latin name (genus species)
- 'common_name': The standardized common name
- 'variety_name': The variety/cultivar name if mentioned

Plant: {plant_text}

If any information cannot be determined, use empty string for that field."""
    
    response = call_openai_with_retry(prompt)
    if response:
        try:
            data = json.loads(response)
            return {
                'latin_name': data.get('latin_name', ''),
                'common_name': data.get('common_name', ''),
                'variety_name': data.get('variety_name', variety_name)
            }
        except json.JSONDecodeError:
            pass
    
    # Fallback to original hardcoded method
    return get_latin_name_fallback(plant_name, variety_name)

def get_latin_name_fallback(plant_name, variety_name=''):
    """Fallback method for getting Latin name based on common plant knowledge"""
    if pd.isna(plant_name) or plant_name == '':
        return {'latin_name': '', 'common_name': '', 'variety_name': variety_name}
    
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
            return {
                'latin_name': latin_name,
                'common_name': common_name.title(),
                'variety_name': variety_name
            }
    
    return {'latin_name': '', 'common_name': plant_name, 'variety_name': variety_name}

def get_latin_name(plant_name):
    """Legacy function for backward compatibility"""
    result = get_plant_info_openai(plant_name)
    return result['latin_name']

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
            # Convert row to dictionary for AI processing
            row_dict = row.to_dict()
            
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
                
                # Use AI to extract all fields from complete row data
                ai_extracted = extract_all_fields_openai(row_dict)
                
                if ai_extracted:
                    # Use AI-extracted data as primary source
                    new_row.update(ai_extracted)
                    
                    # Override variety name with current variety being processed
                    if variety and variety.strip():
                        new_row['Variety Name species'] = variety
                else:
                    # Fallback to original parsing if AI fails
                    print(f"AI extraction failed for row, using fallback methods...")
                    
                    # Basic field mappings - handle flexible column name matching
                    date_received_col = None
                    entry_no_col = None
                    
                    # Find Date Received column (flexible matching)
                    for col in df.columns:
                        if 'date' in col.lower() and ('received' in col.lower() or 'recieved' in col.lower()):
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
                        new_row['DateReceived'] = format_date_received(row[date_received_col])
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
                        plant_info = get_plant_info_openai(plant_name, variety)
                        new_row['Common Name species'] = plant_info['common_name'] or plant_name
                        new_row['Latin Name species'] = plant_info['latin_name']
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
                
                # Always add DateReceived and IDAssigned from original columns if not set by AI
                if not new_row.get('DateReceived'):
                    date_received_col = None
                    for col in df.columns:
                        if 'date' in col.lower() and ('received' in col.lower() or 'recieved' in col.lower()):
                            date_received_col = col
                            break
                    if date_received_col:
                        new_row['DateReceived'] = format_date_received(row[date_received_col])
                
                if not new_row.get('IDAssigned'):
                    entry_no_col = None
                    for col in df.columns:
                        if 'entry' in col.lower() and 'no' in col.lower():
                            entry_no_col = col
                            break
                        elif col.lower().strip() in ['entry no', 'entryno', 'entry_no', 'id', 'entry id']:
                            entry_no_col = col
                            break
                    if entry_no_col:
                        new_row['IDAssigned'] = row[entry_no_col]
                
                # Process dose fields if not handled by AI
                if not any(new_row.get(f'dose {i}') for i in range(1, 11)):
                    dose_col = None
                    for col in df.columns:
                        if 'dose' in col.lower():
                            dose_col = col
                            break
                    if dose_col:
                        dose_data = process_dose_field(row[dose_col])
                        new_row.update(dose_data)
                
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