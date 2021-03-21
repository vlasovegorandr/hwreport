import os
import subprocess
from pathlib import Path
import time
import csv
import re

def ping(host):
    ''' Возвращает True, если до хоста доходят пинги '''
    command = ['ping', '/n', '2', '/w', '2000', host]
    detached_process_flag = 8
    return subprocess.call(command, creationflags=detached_process_flag) == 0   

def delete_software_info(hwreport_file_path):
    ''' Оставляет в репорте только инфу об аппаратной части ПК '''
    software_info_start_line = '[Software Environment]'
    hardwareonly_dir = Path(__file__).parent.joinpath('MsInfo32Reports/hardware_only_reports')
    hardwareonly_dir.mkdir(parents=True, exist_ok=True)
    with open(hwreport_file_path, 'r', encoding='utf-16') as file:
        hardwareonly_file_path = Path.joinpath(hardwareonly_dir, hwreport_file_path.name.split('.')[0]+'_hardware_report.txt')
        with open(hardwareonly_file_path, 'w', encoding='utf-16') as hardware_only_file:
            for line in file:            
                if software_info_start_line in line:
                    break
                hardware_only_file.write(line)
    return hardwareonly_file_path
                
def mark_computer_as_completed(computer_name):
    with open('computer_names.txt', 'r') as file:
        file_contents = file.readlines()
    with open('computer_names.txt', 'w') as file:
        for line in file_contents:
            if computer_name not in line.lower():
                file.write(line)
    with open('completed_computers.txt', 'a') as file:
        file.write(computer_name+'\n')

def parse_file(file_name):
    hardware_info = dict.fromkeys(['Доменное имя пк', 'Процессор', 'Материнская плата', 'Оперативная память', 'Видеокарта', 'Модель диска', 'Размер диска'])
    with open(file_name, 'r', encoding='utf-16') as file:
        file_contents = file.read()
        categories = []
        for idx, category in enumerate(file_contents.split('[')):
            #skipping file header
            if idx != 0:
                categories.append('['+category)
        system_name = ''
        baseboard_info = ''
        processor = ''
        ram = ''
        videocard = ''
        disk_models = []
        disk_sizes = []
        ru_disk_first_category_passed = False
        for category in categories:                        
            for line in category.replace('\t', ' ').split('\n'):
                if category.startswith(('[System Summary]', '[Сведения о системе]')):                    
                    if line.startswith(('System Name ', 'Имя системы ')):
                        system_name = line.replace('System Name ', '').replace('Имя системы ', '').strip()
                    if line.startswith(('BaseBoard Manufacturer ', 'Изготовитель основной платы ')):
                        baseboard_info += line.replace('BaseBoard Manufacturer ', '').replace('Изготовитель основной платы ', '').strip()
                    if line.startswith(('BaseBoard Product ', 'Модель основной платы ')):
                        baseboard_info += line.replace('BaseBoard Product ', '').replace('Модель основной платы ', '').strip()
                    if line.startswith(('Processor ', 'Процессор ')):
                        processor = line.replace('Processor ', '').replace('Процессор ', '').split('@')[0].strip()
                    if line.startswith(('Total Physical Memory ', 'Полный объем физической памяти ')):
                        ram = line.replace('Total Physical Memory ', '').replace('Полный объем физической памяти ', '').strip()
                    hardware_info['Доменное имя пк'] = system_name
                    hardware_info['Процессор'] = processor
                    hardware_info['Материнская плата'] = baseboard_info
                    hardware_info['Оперативная память'] = ram
                if category.startswith(('[Display]', '[Дисплей]')):                    
                    if line.startswith(('Name ', 'Имя ')):
                        videocard = line.replace('Name ', '').replace('Имя', '').strip()
                    hardware_info['Видеокарта'] = videocard
                if category.startswith('[Disks]'):                    
                    if line.startswith('Model '):
                        disk_models.append(line.replace('Model ', '').strip())
                    if line.startswith('Size '):
                        for match in re.findall(r'(\d+[\,\.]\d{2}\s\w{2})', line):
                            disk_sizes.append(match)
                    hardware_info['Модель диска'] = ', '.join(disk_models)
                    hardware_info['Размер диска'] = ', '.join(disk_sizes)
                if category.startswith('[Диски]'):
                    # в рус локализации drives и disks оба переведены как "диски", пришлось выкручиваться
                    if not ru_disk_first_category_passed:                        
                        ru_disk_first_category_passed = True
                        break
                    if line.startswith('Модель '):                        
                        disk_models.append(line.replace('Модель ', '').strip())                        
                    if line.startswith('Размер '):
                        try:
                            disk_size_field_start = re.match(r'Размер\s\d', line.strip())
                            if line.startswith(disk_size_field_start[0]):                      
                                for match in re.findall(r'(\d+[\,\.]\d{2}\s\w{2})', line):
                                    disk_sizes.append(match)                            
                        except TypeError:                            
                            continue
                    hardware_info['Модель диска'] = ', '.join(disk_models)
                    hardware_info['Размер диска'] = ', '.join(disk_sizes)
    
    return hardware_info

def get_csv_summary(hardware_info):
    parsed_info_dir = Path(__file__).parent.joinpath('MsInfo32Reports/hardware_only_reports/summary')
    parsed_info_dir.mkdir(parents=True, exist_ok=True)
    summary_file = parsed_info_dir.joinpath('summary.csv')
    with open(summary_file, 'a', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        if csv_file.tell() == 0:
            csv_writer.writerow(hardware_info.keys())
        csv_writer.writerow(hardware_info.values())

def get_txt_summary(hardware_info):
    parsed_info_dir = Path(__file__).parent.joinpath('MsInfo32Reports/hardware_only_reports/summary')
    parsed_info_dir.mkdir(parents=True, exist_ok=True)
    summary_file = parsed_info_dir.joinpath('summary.txt')
    with open(summary_file, 'a') as file:
        for key, item in hardware_info.items():
            if type(item) is list:
                item = ', '.join(item)
            file.write(key+': '+item+'\n')
        file.write('\n')       
            
def create_reports():
    hwreports_dir = Path(__file__).parent.joinpath('MsInfo32Reports')
    hwreports_dir.mkdir(parents=True, exist_ok=True)
    try:          
        with open('computer_names.txt', 'r') as file:
            for computer_name in file:
                computer_name = computer_name.lower().strip('\n ')
                if not computer_name:
                    continue
                report_path = hwreports_dir.joinpath(computer_name+".txt")
                if ping(computer_name):
                    print(f'{time.strftime("%H:%M:%S")} - started getting info for {computer_name}...')
                    os.system(f'cmd /c "msinfo32 /computer {computer_name} /report {report_path}')
                    print(f'{time.strftime("%H:%M:%S")} - report completed for {computer_name}...')
                    hardware_only_file = delete_software_info(report_path)
                    print(f'{time.strftime("%H:%M:%S")} - hardware-only report created for {computer_name}...')
                    hardware_info = parse_file(hardware_only_file)
                    get_txt_summary(hardware_info)
                    get_csv_summary(hardware_info)
                    print(f'{time.strftime("%H:%M:%S")} - {computer_name} info added to summary...')
                    mark_computer_as_completed(computer_name)
                    print('\n')
    except FileNotFoundError:        
        with open('computer_names.txt', 'w') as file:
            file.write('Укажите доменные имена (или IP-адреса) компьютеров\n')
            file.write('для сбора информации,\n')
            file.write('один компьютер - одна строка.\n')
            file.write('Пример:\n')
            file.write('Компьютер 1\n')
            file.write('Компьютер 2\n')
            file.write('Компьютер 3\n')
        print('\nФайл с именами компьютеров не найден, создали новый\n')

if __name__ == '__main__':
    create_reports()