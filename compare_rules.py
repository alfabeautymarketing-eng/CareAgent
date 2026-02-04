import json

with open('remote_rules.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("Rules targeting 'Заказ':")
rules_zakaz = []
for rule in data['rules']:
    if rule.get('target_sheet') == 'Заказ':
        rules_zakaz.append(rule)
        print(f"ID: {rule['id']} | {rule['source_sheet']} -> {rule['target_sheet']} ({rule['target_header']})")

print("\nRules targeting 'Заказ2026':")
rules_2026 = []
for rule in data['rules']:
    if rule.get('target_sheet') == 'Заказ2026':
        rules_2026.append(rule)
        print(f"ID: {rule['id']} | {rule['source_sheet']} -> {rule['target_sheet']} ({rule['target_header']})")
