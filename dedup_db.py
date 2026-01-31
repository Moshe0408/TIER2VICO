import json

def deduplicate_db():
    try:
        with open('integrations_db.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        seen = set()
        unique_data = []
        for item in data:
            name = item.get('Customer', '').strip()
            if name and name not in seen:
                seen.add(name)
                # Ensure Volume is removed as per previous request
                if 'Volume' in item:
                    del item['Volume']
                # Status is still in JSON for logic but will be hidden from UI
                unique_data.append(item)
        
        with open('integrations_db.json', 'w', encoding='utf-8') as f:
            json.dump(unique_data, f, ensure_ascii=False, indent=4)
        print(f"Deduplicated database. Now has {len(unique_data)} records.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    deduplicate_db()
