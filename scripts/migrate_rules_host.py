#!/usr/bin/env python3
import sys
import os
import yaml
import json
from pathlib import Path

# Try to import gspread, if not installed, we can't run.
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    print("❌ Libraries missing. Please run: pip3 install gspread google-auth PyYAML")
    sys.exit(1)

# Configuration
BASE_DIR = Path("/root/AgentCare")
CREDENTIALS_PATH = BASE_DIR / "config/credentials.json"
RULES_DIR = BASE_DIR / "config/rules"

# Project Mapping (Spreadsheet IDs)
PROJECT_MAP = {
    "MT": "13kB77R67GJOZQ3vsLcwR1nUaRsupR8ZnEaTdDd66CTQ",
    "SS": "12yIL1CuESZxeUUd-oKK2brtN1FnXE9q95N7SqzNc7vk",
    "SK": "1CpYYLvRYslsyCkuLzL9EbbjsvbNpWCEZcmhKqMoX5zw"
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def connect_to_sheets():
    if not CREDENTIALS_PATH.exists():
        print(f"❌ Credentials not found at {CREDENTIALS_PATH}")
        sys.exit(1)
    
    creds = Credentials.from_service_account_file(
        str(CREDENTIALS_PATH), 
        scopes=SCOPES
    )
    return gspread.authorize(creds)

def migrate_rules_for_project(gc, name, ss_id):
    print(f"🚀 Migrating rules for {name} ({ss_id})...")
    
    try:
        sh = gc.open_by_key(ss_id)
        # Check if sheet exists
        try:
            ws = sh.worksheet("Правила синхро")
        except gspread.WorksheetNotFound:
            print(f"⚠️  Sheet 'Правила синхро' not found in {name}. Skipping.")
            return

        values = ws.get_all_values()
        if not values or len(values) < 2:
            print(f"⚠️  Sheet 'Правила синхро' is empty in {name}. Skipping.")
            return

        headers = [str(h).strip().lower() for h in values[0]]
        print(f"DEBUG Headers found in {name}: {headers}")
        
        # Helper to get value
        def get_val(row, possible_names):
            possible_names = [n.lower() for n in possible_names]
            for i, h in enumerate(headers):
                if h in possible_names:
                    if i < len(row):
                        return str(row[i]).strip()
            return ""

        rules = []
        
        # Process rows
        for i, row in enumerate(values[1:]):
            if not any(row): continue
            
            # Determine mode
            mode = 'unidirectional'
            raw_mode = get_val(row, ['mode', 'режим', 'тип'])
            if 'bi' in raw_mode.lower() or 'дву' in raw_mode.lower():
                mode = 'bidirectional'

            rule = {
                'id': f"rule_{i+1}", 
                'mode': mode,
                'enabled': True,
                'category': 'Миграция',
            }

            # Check enabled
            enabled_val = get_val(row, ['enabled', 'активно', 'включено', 'status', 'active'])
            if enabled_val and enabled_val.lower() in ['false', '0', 'no', 'нет']:
                rule['enabled'] = False

            # Load metadata
            rule['category'] = get_val(row, ['category', 'категория']) or rule['category']

            if mode == 'bidirectional':
                rule['sheet_a'] = get_val(row, ['sheet a', 'лист а', 'лист 1', 'source sheet'])
                rule['header_a'] = get_val(row, ['header a', 'колонку а', 'колонка 1', 'source header'])
                rule['sheet_b'] = get_val(row, ['sheet b', 'лист б', 'лист 2', 'target sheet'])
                rule['header_b'] = get_val(row, ['header b', 'колонку б', 'колонка 2', 'target header'])
            else:
                rule['source_sheet'] = get_val(row, ['source sheet', 'источник лист', 'откуда', 'источник: лист', 'исходный лист'])
                rule['source_header'] = get_val(row, ['source header', 'источник колонка', 'что', 'источник: колонка', 'заголовок исх. столбца'])
                rule['target_sheet'] = get_val(row, ['target sheet', 'цель лист', 'куда', 'цель: лист', 'целевой лист'])
                rule['target_header'] = get_val(row, ['target header', 'цель колонка', 'куда колонка', 'цель: колонка', 'заголовок цел. столбца'])

            # Validate
            valid = False
            if mode == 'bidirectional':
                if rule.get('sheet_a') and rule.get('sheet_b'): valid = True
            else:
                if rule.get('source_sheet') and rule.get('target_sheet'): valid = True
            
            if valid:
                import uuid
                rule['id'] = str(uuid.uuid4())
                rules.append(rule)

        if not rules:
            print(f"ℹ️  No valid rules found in {name}.")
            return

        # Save to YAML
        config_path = RULES_DIR / f"{ss_id}.yaml"
        # Ensure dir exists
        RULES_DIR.mkdir(parents=True, exist_ok=True)
        
        final_rules = rules
        if config_path.exists():
             try:
                with open(config_path, 'r') as f:
                    existing = yaml.safe_load(f) or []
                    print(f"ℹ️  Existing config found with {len(existing)} rules. Appending migrated rules.")
                    final_rules = existing + rules
             except Exception as e:
                 print(f"⚠️ Error reading existing config: {e}. Overwriting.")

        with open(config_path, 'w') as f:
            yaml.dump(final_rules, f, allow_unicode=True, default_flow_style=False)
            
        print(f"✅ Saved {len(rules)} migrated rules to {config_path}")

        # Delete legacy sheet
        try:
            sh.del_worksheet(ws)
            print("🗑️  Legacy sheet deleted.")
        except Exception as e:
            print(f"⚠️  Could not delete sheet: {e}")

    except Exception as e:
        print(f"❌ Error migrating {name}: {e}")

if __name__ == "__main__":
    print("Starting HOST-BASED Server-Side Rule Migration...")
    gc = connect_to_sheets()
    print("✅ Connected to Google Sheets API")
    
    for name, ssid in PROJECT_MAP.items():
        migrate_rules_for_project(gc, name, ssid)
    print("Done.")
