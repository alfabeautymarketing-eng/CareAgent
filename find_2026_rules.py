import json

with open('remote_rules.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("Rules targeting 'Заказ2026':")
found = False
for rule in data['rules']:
    if rule.get('target_sheet') == 'Заказ2026':
        print(f"ID: {rule['id']} | {rule['source_sheet']} -> {rule['target_sheet']} ({rule['target_header']})")
        found = True

if not found:
    print("None found (explicit match)")
    print("Searching for substring 'Заказ2026' in any field...")
    for rule in data['rules']:
        if 'Заказ2026' in str(rule):
             print(f"ID: {rule['id']} | Match found in fields")
