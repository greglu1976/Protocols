import os
import subprocess

# Ищем все папки, которые начинаются на "ШЭТ"
cabinets = [d for d in os.listdir('.') if os.path.isdir(d) and d.startswith("ШЭТ")]

if not cabinets:
    print("Не найдено папок, начинающихся на 'ШЭТ'!")
    exit()

print(f"Найдены шкафы для сборки: {', '.join(cabinets)}")

for cab in cabinets:
    print(f"\n{'='*40}")
    print(f"Запуск сборки для: {cab}")
    print(f"{'='*40}")
    
    # Передаем имя шкафа в Makefile
    subprocess.run(["make", f"CABINET={cab}"])

print("\nСборка всех шкафов завершена!")