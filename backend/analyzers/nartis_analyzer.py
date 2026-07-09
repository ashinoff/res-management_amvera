#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import sys
from collections import defaultdict
import re
import xlrd
from datetime import datetime

class NartisAnalyzer:
    def __init__(self):
        self.ru_months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 
                          'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']
        self.UNDERVOLTAGE_THRESHOLD = 198
        self.OVERVOLTAGE_THRESHOLD = 242
    
    def analyze_file(self, filepath):
        """Анализ файла журнала событий Нартис"""
        try:
            events_data = {
                'overvoltage': {'A': [], 'B': [], 'C': []},
                'undervoltage': {'A': [], 'B': [], 'C': []}
            }
            
            # ВАЖНО: для .xls файлов используем formatting_info=True
            try:
                # Пробуем с formatting_info для .xls
                workbook = xlrd.open_workbook(filepath, formatting_info=True)
            except:
                # Если не получилось (например .xlsx), открываем обычно
                workbook = xlrd.open_workbook(filepath)
                
            sheet = workbook.sheet_by_index(0)
            
            print(f"Sheet rows: {sheet.nrows}, cols: {sheet.ncols}", file=sys.stderr)
            
            # Проверяем объединенные ячейки
            start_row = 1  # по умолчанию начинаем со второй строки
            
            if hasattr(sheet, 'merged_cells') and len(sheet.merged_cells) > 0:
                print(f"MERGED CELLS FOUND: {sheet.merged_cells}", file=sys.stderr)
                # Если есть объединение в первой строке - начинаем после него
                for (rlo, rhi, clo, chi) in sheet.merged_cells:
                    if rlo == 0:  # объединение начинается с первой строки
                        print(f"Merged cell in first row detected, skipping to row {rhi}", file=sys.stderr)
                        start_row = max(start_row, rhi)  # начинаем после объединенных строк
                        
            print(f"Starting from row: {start_row}", file=sys.stderr)
            
            # Выведем первые 5 строк для проверки
            print("First 5 data rows:", file=sys.stderr)
            for i in range(start_row, min(start_row + 5, sheet.nrows)):
                row = []
                for j in range(sheet.ncols):
                    row.append(str(sheet.cell_value(i, j))[:30])
                print(f"Row {i}: {row}", file=sys.stderr)
            
            # Парсим каждую строку
            for row_idx in range(start_row, sheet.nrows):
                try:
                    datetime_str = str(sheet.cell_value(row_idx, 0)) if sheet.cell_value(row_idx, 0) else ""
                    event = str(sheet.cell_value(row_idx, 1)) if sheet.cell_value(row_idx, 1) else ""
                    
                    if not datetime_str or not event or datetime_str == '0':
                        continue
                    if datetime_str == 'Время' or event == 'Событие журнала напряжений':
                        continue
                    
                    voltage_str = str(sheet.cell_value(row_idx, 2)).replace(',', '.') if sheet.cell_value(row_idx, 2) else "0"
                    
                    try:
                        voltage_raw = float(voltage_str)
                        voltage = voltage_raw / 10.0
                    except ValueError:
                        print(f"Row {row_idx}: cannot parse voltage '{voltage_str}', skipping", file=sys.stderr)
                        continue
            
                    percent_str = str(sheet.cell_value(row_idx, 3)).replace(',', '.') if sheet.cell_value(row_idx, 3) else "0"
                    duration_str = str(sheet.cell_value(row_idx, 4)).replace(',', '.') if sheet.cell_value(row_idx, 4) else "0"
                    
                    if not voltage_str or not percent_str or not duration_str:
                        print(f"Row {row_idx}: empty values, skipping", file=sys.stderr)
                        continue
                    
                    percent = float(percent_str)
                    duration = float(duration_str)
                    
                    if duration <= 60:
                        print(f"Skipped: duration {duration} <= 60", file=sys.stderr)
                        continue
                    
                    if abs(voltage - 11.50) < 0.001 or voltage == 0:
                        print(f"Skipped: voltage {voltage} is 11.50 or 0", file=sys.stderr)
                        continue
                    
                    date_match = re.match(r'(\d{2})\.(\d{2})\.(\d{4})', datetime_str)
                    if date_match:
                        month = int(date_match.group(2))
                    else:
                        print(f"Could not parse date from: {datetime_str}", file=sys.stderr)
                        continue
                    
                    phase = None
                    event_type = None
                    
                    if 'фаза A' in event:
                        phase = 'A'
                    elif 'фаза B' in event:
                        phase = 'B'
                    elif 'фаза C' in event:
                        phase = 'C'
                    
                    if phase:
                        if 'Окончание провала' in event:
                            event_type = 'undervoltage'
                            if voltage >= self.UNDERVOLTAGE_THRESHOLD:
                                continue
                        elif 'Окончание перенапряжения' in event:
                            event_type = 'overvoltage'
                            if voltage <= self.OVERVOLTAGE_THRESHOLD:
                                continue
                    
                    if phase and event_type:
                        events_data[event_type][phase].append({
                            'voltage': voltage,
                            'month': month,
                            'duration': duration
                        })
                        
                except Exception as e:
                    print(f"Error in row {row_idx}: {str(e)}", file=sys.stderr)
                    continue
            
            return self._generate_result(events_data)
            
        except xlrd.biffh.XLRDError as e:
            return {
                'success': False,
                'error': f"Ошибка чтения Excel файла: {str(e)}",
                'has_errors': False
            }
        except Exception as e:
            return {
                'success': False,
                'error': f"Ошибка анализа файла: {str(e)}",
                'has_errors': False
            }
    
    def _generate_result(self, events_data):
        """Генерация результата анализа"""
        summary_parts = []
        has_errors = False
        details = {
            'overvoltage': {},
            'undervoltage': {}
        }
        
        # Обработка перенапряжений
        for phase in ['A', 'B', 'C']:
            events = events_data['overvoltage'][phase]
            if len(events) > 10:
                has_errors = True
                months = [e['month'] for e in events]
                min_month = min(months)
                max_month = max(months)
                
                if min_month == max_month:
                    period = self.ru_months[min_month-1]
                else:
                    period = f"{self.ru_months[min_month-1]}-{self.ru_months[max_month-1]}"
                
                voltages = [e['voltage'] for e in events]
                max_voltage = max(voltages)
                min_voltage_in_overvoltage = min(voltages)
                count = len(events)
                
                # Расчет процентов для диапазона
                min_percent = ((min_voltage_in_overvoltage - 220) / 220) * 100
                max_percent = ((max_voltage - 220) / 220) * 100
                
                summary_parts.append(
                    f"Фаза {phase}: Перенапряжение {min_percent:.1f}-{max_percent:.1f}% (max {max_voltage:.0f}В) ({period}) - {count} событий"
                )
                
                details['overvoltage'][f'phase_{phase}'] = {
                    'count': count,
                    'max': max_voltage,
                    'period': period
                }
        
        # Обработка провалов
        for phase in ['A', 'B', 'C']:
            events = events_data['undervoltage'][phase]
            if len(events) > 10:
                has_errors = True
                months = [e['month'] for e in events]
                min_month = min(months)
                max_month = max(months)
                
                if min_month == max_month:
                    period = self.ru_months[min_month-1]
                else:
                    period = f"{self.ru_months[min_month-1]}-{self.ru_months[max_month-1]}"
                
                voltages = [e['voltage'] for e in events]
                min_voltage = min(voltages)
                max_voltage_in_undervoltage = max(voltages)
                count = len(events)
                
                # Расчет процентов для диапазона
                min_percent = ((220 - max_voltage_in_undervoltage) / 220) * 100
                max_percent = ((220 - min_voltage) / 220) * 100
                
                summary_parts.append(
                    f"Фаза {phase}: Провал {min_percent:.1f}-{max_percent:.1f}% (min {min_voltage:.0f}В) ({period}) - {count} событий"
                )
                
                details['undervoltage'][f'phase_{phase}'] = {
                    'count': count,
                    'min': min_voltage,
                    'period': period
                }
        
        if not has_errors:
            total_events = sum(len(events) for events in events_data['overvoltage'].values())
            total_events += sum(len(events) for events in events_data['undervoltage'].values())
            
            if total_events > 0:
                summary = f"Обнаружено событий: {total_events}, но все менее 10 по каждому типу"
            else:
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
    
    try:
        analyzer = NartisAnalyzer()
        result = analyzer.analyze_file(sys.argv[1])
        print(json.dumps(result, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({'success': False, 'error': str(e)}))
