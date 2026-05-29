import sys
import shutil
import pandas as pd
from pathlib import Path

# Пути можно передать аргументами или зашить
qc_xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r'C:\Users\Lab\Desktop\Allab\ALLAB_2026\2026\qc_data.xlsx')
target  = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(r'C:\Users\Lab\Desktop\for alibek')

target.mkdir(parents=True, exist_ok=True)

df = pd.read_excel(qc_xlsx, sheet_name='raw_sheet', dtype=str)
paths = df['file_name'].dropna().str.strip().tolist()

# Уникализируем — мало ли один и тот же файл попал в raw_sheet дважды
paths = list(dict.fromkeys(paths))

ok = missing = errors = 0
for p in paths:
    src = Path(p)
    if not src.is_file():
        print(f'НЕТ: {p}')
        missing += 1
        continue
    try:
        dst = target / src.name
        # Если одинаковые имена из разных папок — добавим суффикс
        if dst.exists() and dst.resolve() != src.resolve():
            i = 1
            while True:
                cand = target / f'{src.stem}__{i}{src.suffix}'
                if not cand.exists():
                    dst = cand
                    break
                i += 1
        shutil.copy2(src, dst)
        print(f'OK : {src.name}')
        ok += 1
    except Exception as e:
        print(f'ERR: {p} — {e}')
        errors += 1

print(f'\nИтого: скопировано {ok}, не найдено {missing}, ошибок {errors}')