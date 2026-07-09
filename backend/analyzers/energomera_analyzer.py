#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import sys
import csv
import pandas as pd
import re
from collections import defaultdict

class EnergomeraAnalyzer:
    def __init__(self):
        self.ru_months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 
                          'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']
    
    def analyze_file(self, filepath):
        """Анализ файла журнала событий Энергомера"""
        try:
            # Читаем Excel файл, пропуская первые 15 строк (заголовки с 16-й)
            df = pd.read_excel(filepath, header=None, skiprows=15)
            
            # Сохраняем во временный CSV для обработки
            temp_csv = filepath + '.temp.csv'
            df.to_csv(temp_csv, index=False, header=False)
            
            # Читаем CSV
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
            
            # Пропускаем строку с заголовками и обрабатываем данные
            # Заголовки: Дата/время | Событие | Порог напряжения, В | Порог, % | Мин./макс. значение напряжения, В | Глубина/высота/уровень, % | Длительность, с | Время работы счетчика
            for idx in range(1, len(rows)):  # Пропускаем строку заголовков
                try:
                    row = rows[idx]
                    if len(row) < 7:  # Проверяем минимальное количество колонок
                        continue
                    
                    # Извлекаем данные из колонок
                    datetime_str = row[0]  # A - Дата/время
                    event = row[1]         # B - Событие
                    # C - Порог напряжения (не используем)
                    # D - Порог, % (не используем)
                    voltage_str = row[4]   # E - Мин./макс. значение напряжения, В
                    # F - Глубина/высота/уровень, % (не используем)
                    duration_str = row[6]  # G - Длительность, с
                    
                    # Парсим значения
                    voltage = float(voltage_str.replace(',', '.'))
                    duration = float(duration_str.replace(',', '.'))
                    
                    # Фильтры
                    if duration <= 60:
                        continue
                    if abs(voltage - 11.50) < 0.001 or voltage == 0:
                        continue
                    
                    # Извлекаем месяц из даты (формат: 08.08.2025 07:34:16)
                    date_match = re.match(r'(\d{2})\.(\d{2})\.(\d{4})', datetime_str)
                    if not date_match:
                        continue
                    month = int(date_match.group(2))
                    
                    # Определяем фазу и тип события
                    phase = None
                    event_type = None
                    event_lower = event.lower()
                    
                    # Определяем фазу (Фаза А, Фаза B, Фаза С)
                    if 'фаза а' in event_lower:
                        phase = 'A'
                    elif 'фаза b' in event_lower:
                        phase = 'B'
                    elif 'фаза с' in event_lower or 'фаза c' in event_lower:
                        phase = 'C'
                    
                    # Определяем тип события
                    if phase and 'окончание' in event_lower:
                        if 'провал' in event_lower:
                            event_type = 'undervoltage'
                        elif 'перенапряжение' in event_lower:
                            event_type = 'overvoltage'
                    
                    # Сохраняем событие
                    if phase and event_type:
                        events_data[event_type][phase].append({
                            'voltage': voltage,
                            'month': month,
                            'duration': duration
                        })
                        
                except Exception as e:
                    continue
            
            return self._generate_result(events_data)
            
        except Exception as e:
            return {
                'success': False,
                'error': f"Ошибка анализа: {str(e)}",
                'has_errors': False
            }
    
    def _generate_result(self, events_data):
        """Генерация результата анализа"""
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
        
        return {
            'success': True,
            'summary': summary,
            'has_errors': has_errors,
            'details': details
        }

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({'success': False, 'error': 'No file path provided'}))
        sys.exit(1)
    
    analyzer = EnergomeraAnalyzer()
    result = analyzer.analyze_file(sys.argv[1])
    print(json.dumps(result, ensure_ascii=False))
