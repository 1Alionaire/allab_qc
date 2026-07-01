import json

def extract_unique_file_names(input_path: str, output_path: str):
    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    unique_files = sorted(set(
        entry["file_name"]
        for entry in data
        if "file_name" in entry
    ))

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(unique_files, f, ensure_ascii=False, indent=2)

    print(f"Найдено уникальных файлов: {len(unique_files)}")
    print(f"Сохранено в: {output_path}")


if __name__ == "__main__":
    extract_unique_file_names("total_data.json", "all_files.json")