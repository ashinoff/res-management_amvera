#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import sys
import csv
import pandas as pd
import re
from collections import defaultdict

class RIMAnalyzer:
    def __init__(self):
        self.ru_months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 
                          'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']
    
    def analyze_file(self, filepath):
        """Анализ файла журнала событий через промежуточный CSV"""
        try:
            # Читаем весь Excel файл без пропуска строк
            df = pd.read_excel(filepath, header=None)
            
            # Сохраняем во временный CSV
            temp_csv = filepath + '.temp.csv'
            df.to_csv(temp_csv, index=False, header=False)
            
            # Теперь читаем CSV
            events_data = {
                'overvoltage': {'A': [], 'B': [], 'C': []},
                'undervoltage': {'A': [], 'B': [], 'C': []}
            }
            
            with open(temp_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
            
            # Удаляем временный файл
            import os
            os.remove(temp_csv)
            
            # ОТЛАДКА: Выводим первые несколько строк
            print("=== ПЕРВЫЕ 5 СТРОК ФАЙЛА ===", file=sys.stderr)
            for i in range(min(5, len(rows))):
                print(f"Строка {i}: {rows[i][:5] if len(rows[i]) > 5 else rows[i]}", file=sys.stderr)
            
            # Ищем строку со словом "Время" в первой колонке
            data_start_row = 0
            for idx, row in enumerate(rows):
                if len(row) >= 1:
                    # Проверяем только первую колонку на наличие слова "Время"
                    if 'время' in str(row[0]).lower():
                        print(f"=== Нашли заголовок 'Время' в строке {idx} ===", file=sys.stderr)
                        data_start_row = idx + 1  # Данные начинаются со следующей строки
                        break
            
            print(f"=== Начинаем обработку с строки {data_start_row} ===", file=sys.stderr)
            
            # Счетчики для отладки
            total_rows = 0
            total_events = 0
            filtered_by_duration = 0
            filtered_by_voltage = 0
            no_date = 0
            no_phase = 0
            
            # Обрабатываем данные начиная с найденной строки
            for idx in range(data_start_row, len(rows)):
                try:
                    row = rows[idx]
                    total_rows += 1
                    
                    if len(row) < 5:
                        continue
                    
                    # Фиксированная структура колонок:
                    # A(0) - Время, B(1) - Событие, C(2) - Напряжение, D(3) - %, E(4) - Продолжительность
                    datetime_str = str(row[0])
                    event = str(row[1])
                    
                    # Парсим числа с запятой
                    try:
                        voltage = float(str(row[2]).replace(',', '.'))
                    except:
                        continue
                        
                    try:
                        duration = float(str(row[4]).replace(',', '.'))
                    except:
                        continue
                    
                    # Фильтры
                    if duration <= 60:
                        filtered_by_duration += 1
                        continue
                    if abs(voltage - 11.50) < 0.001 or voltage == 0:
                        filtered_by_voltage += 1
                        continue
                    
                    # Извлекаем месяц из даты
                    date_match = re.match(r'(\d{2})\.(\d{2})\.(\d{4})', datetime_str)
                    if not date_match:
                        no_date += 1
                        continue
                    month = int(date_match.group(2))
                    
                    # Определяем фазу и тип события
                    phase = None
                    event_type = None
                    event_lower = event.lower()
                    
                    if 'фаза a' in event_lower:
                        phase = 'A'
                    elif 'фаза b' in event_lower:
                        phase = 'B'
                    elif 'фаза c' in event_lower:
                        phase = 'C'
                    
                    if phase and 'окончание' in event_lower:
                        if 'провал' in event_lower:
                            event_type = 'undervoltage'
                        elif 'перенапряжение' in event_lower:
                            event_type = 'overvoltage'
                    
                    if phase and event_type:
                        total_events += 1
                        events_data[event_type][phase].append({
                            'voltage': voltage,
                            'month': month,
                            'duration': duration
                        })
                        # Выводим первые несколько найденных событий
                        if total_events <= 3:
                            print(f"=== Найдено событие: {event}, V={voltage}, T={duration}с ===", file=sys.stderr)
                    else:
                        no_phase += 1
                        
                except Exception as e:
                    print(f"Ошибка в строке {idx}: {e}", file=sys.stderr)
                    continue
            
            # Отладочная статистика
            print(f"\n=== СТАТИСТИКА ОБРАБОТКИ ===", file=sys.stderr)
            print(f"Всего строк обработано: {total_rows}", file=sys.stderr)
            print(f"Событий найдено: {total_events}", file=sys.stderr)
            print(f"Отфильтровано по длительности (<=60с): {filtered_by_duration}", file=sys.stderr)
            print(f"Отфильтровано по напряжению (11.5В или 0В): {filtered_by_voltage}", file=sys.stderr)
            print(f"Не найдена дата: {no_date}", file=sys.stderr)
            print(f"Не определена фаза/тип: {no_phase}", file=sys.stderr)
            
            # Выводим количество событий по типам
            for event_type in ['overvoltage', 'undervoltage']:
                for phase in ['A', 'B', 'C']:
                    count = len(events_data[event_type][phase])
                    if count > 0:
                        print(f"{event_type} фаза {phase}: {count} событий", file=sys.stderr)
            
            return self._generate_result(events_data)
            
        except Exception as e:
            return {
                'success': False,
                'error': f"Ошибка анализа: {str(e)}",
                'has_errors': False
            }
    
    def _generate_result(self, events_data):
        summary_parts = []
        has_errors = False
        details = {'overvoltage': {}, 'undervoltage': {}}
        
        # Обработка перенапряжений
        for phase in ['A', 'B', 'C']:
            events = events_data['overvoltage'][phase]
            if len(events) > 10:
                has_errors = True
                months = [e['month'] for e in events]
                min_month, max_month = min(months), max(months)
                period = self.ru_months[min_month-1] if min_month == max_month else f"{self.ru_months[min_month-1]}-{self.ru_months[max_month-1]}"
                voltages = [e['voltage'] for e in events]
                max_voltage = max(voltages)
                min_voltage_in_overvoltage = min(voltages)
                count = len(events)
                
                # Расчет процентов для диапазона
                min_percent = ((min_voltage_in_overvoltage - 220) / 220) * 100
                max_percent = ((max_voltage - 220) / 220) * 100
                
                summary_parts.append(f"Фаза {phase}: Перенапряжение {min_percent:.1f}-{max_percent:.1f}% (max {max_voltage:.0f}В) ({period}) - {count} событий")
                details['overvoltage'][f'phase_{phase}'] = {'count': count, 'max': max_voltage, 'period': period}
        
        # Обработка провалов
        for phase in ['A', 'B', 'C']:
            events = events_data['undervoltage'][phase]
            if len(events) > 10:
                has_errors = True
                months = [e['month'] for e in events]
                min_month, max_month = min(months), max(months)
                period = self.ru_months[min_month-1] if min_month == max_month else f"{self.ru_months[min_month-1]}-{self.ru_months[max_month-1]}"
                voltages = [e['voltage'] for e in events]
                min_voltage = min(voltages)
                max_voltage_in_undervoltage = max(voltages)
                count = len(events)
                
                # Расчет процентов для диапазона
                min_percent = ((220 - max_voltage_in_undervoltage) / 220) * 100
                max_percent = ((220 - min_voltage) / 220) * 100
                
                summary_parts.append(f"Фаза {phase}: Провал {min_percent:.1f}-{max_percent:.1f}% (min {min_voltage:.0f}В) ({period}) - {count} событий")
                details['undervoltage'][f'phase_{phase}'] = {'count': count, 'min': min_voltage, 'period': period}
        
        if not has_errors:
            summary = "Напряжение в пределах ГОСТ"
        else:
            summary = '; '.join(summary_parts)
        
        print(f"\n=== РЕЗУЛЬТАТ: {summary} ===", file=sys.stderr)
        
        return {'success': True, 'summary': summary, 'has_errors': has_errors, 'details': details}

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({'success': False, 'error': 'No file path provided'}))
        sys.exit(1)
    
    analyzer = RIMAnalyzer()
    result = analyzer.analyze_file(sys.argv[1])
    print(json.dumps(result, ensure_ascii=False))
