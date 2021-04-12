import sys
import os
import subprocess
from pathlib import Path
import time
import csv
import re
import logging
from logging import StreamHandler, Formatter

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = StreamHandler(stream=sys.stdout)
handler.setFormatter(Formatter('[%(asctime)s] %(message)s', '%H:%M:%S'))
logger.addHandler(handler)

def ping(host):
    ''' Возвращает True, если хост отвечает на пинги '''
    command = ['ping', '/n', '2', '/w', '2000', host]
    detached_process_flag = 8
    return subprocess.call(command, creationflags=detached_process_flag) == 0   

def create_msinfo32_report(computer_name, report_path):
    os.system(f'cmd /c "msinfo32 /computer {computer_name} /report {report_path}')

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

def get_csv_summary(hardware_info, script_start_datetime):
    parsed_info_dir = Path(__file__).parent.joinpath('MsInfo32Reports/hardware_only_reports/summary')
    parsed_info_dir.mkdir(parents=True, exist_ok=True)
    # В имя файла вставлять дату
    summary_file = parsed_info_dir.joinpath(f'summary-{script_start_datetime}.csv')
    with open(summary_file, 'a', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        if csv_file.tell() == 0:
            csv_writer.writerow(hardware_info.keys())
        csv_writer.writerow(hardware_info.values())

def get_txt_summary(hardware_info, script_start_datetime):
    parsed_info_dir = Path(__file__).parent.joinpath('MsInfo32Reports/hardware_only_reports/summary')
    parsed_info_dir.mkdir(parents=True, exist_ok=True)
    summary_file = parsed_info_dir.joinpath(f'summary-{script_start_datetime}.txt')
    with open(summary_file, 'a') as file:
        for key, item in hardware_info.items():
            if type(item) is list:
                item = ', '.join(item)
            file.write(key+': '+item+'\n')
        file.write('\n')       
            
def create_reports():
    hwreports_dir = Path(__file__).parent.joinpath('MsInfo32Reports')
    hwreports_dir.mkdir(parents=True, exist_ok=True)
    script_start_datetime = time.strftime("%H_%M_%S-%d_%m_%y")
    try:
        print('\nВ результате работы программы создаются следующие папки:')
        print('- MsInfo32Reports - содержит полные отчеты по компьютерам')
        print('- MsInfo32Reports/hardware_only_reports - отчеты только по аппаратной части компьютеров')
        print('- MsInfo32Reports/hardware_only_reports/summary - общие отчеты в txt и csv, содержащие данные о процессоре,')     
        print('    материнке, оперативке, видеокарте и дисках всех компов')
        input('---Нажмите Enter для продолжения---\n')
        with open('computer_names.txt', 'r', encoding='utf-8') as file:
            done_computers = []
            failed_computers = []                        
            for computer_name in file:
                computer_name = computer_name.lower().strip('\n ')
                if not computer_name:
                    continue
                report_path = hwreports_dir.joinpath(computer_name+".txt")
                if ping(computer_name):
                    logger.info(f'начали собирать инфу по компьютеру {computer_name}...')
                    create_msinfo32_report(computer_name, report_path)
                    logger.info(f'отчет создан по компьютеру {computer_name}...')
                    hardware_only_file = delete_software_info(report_path)
                    hardware_info = parse_file(hardware_only_file)
                    get_txt_summary(hardware_info, script_start_datetime)
                    get_csv_summary(hardware_info, script_start_datetime)                    
                    logger.info(f'инфа по {computer_name} добавлена в общий отчет...')
                    done_computers.append(computer_name)
                    print('\n')
                else:
                    failed_computers.append(computer_name)
            print('---Итоги---')
            if done_computers:
                print('\nСобрали инфу по компьютерам:')
                for computer_name in done_computers:
                    print(' - '+computer_name)
            if failed_computers:
                print('\nНе удалось собрать инфу по компьютерам:')
                for computer_name in failed_computers:
                    print(' - '+computer_name)
    except FileNotFoundError:        
        with open('computer_names.txt', 'w', encoding='utf-8') as file:
            file.write('Компьютер 1\n')
            file.write('Компьютер 2\n')
            file.write('Компьютер 3\n')
        print('\nФайл с именами компьютеров не найден, создали новый.\n')
        print('Укажите в этом файле доменные имена (или IP-адреса) компьютеров')
        print('  для сбора информации, один компьютер - одна строка.')
        print('Все данные, кроме имен пк, нужно удалить из файла.')
        print('После этого запустите программу еще раз')
        input('---Нажмите Enter для продолжения---')

if __name__ == '__main__':
    create_reports()