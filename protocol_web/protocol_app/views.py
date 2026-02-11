from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.contrib import messages
from django.shortcuts import redirect
from django.views.decorators.http import require_http_methods
from django.core.paginator import Paginator
from django.db.models import Q
from django.conf import settings
import os
from datetime import datetime

from django.views.decorators.csrf import csrf_exempt

from .models import Protocol
from .forms import ProtocolUploadForm, ProtocolSearchForm, ProtocolExportForm, MultipleFileField

import logging
from sqlalchemy import create_engine
import pandas as pd
import xlsxwriter



logger = logging.getLogger(__name__)

# Временно упрощенные представления для создания миграций

def index(request):
    """Главная страница"""
    total_protocols = 0
    pel_count = 0
    ai_count = 0

    try:
        total_protocols = Protocol.objects.count()
        pel_count = Protocol.objects.filter(type_protocol__startswith='pel').count()
        ai_count = Protocol.objects.filter(type_protocol__startswith='ai').count()
    except:
        pass

    context = {
        'total_protocols': total_protocols,
        'pel_count': pel_count,
        'ai_count': ai_count,
    }
    return render(request, 'protocol_app/index.html', context)


@require_http_methods(['GET', 'POST'])
def upload_protocols(request):
    """Загрузка и обработка протоколов"""
    if request.method == 'POST':
        form = ProtocolUploadForm(request.POST, request.FILES)
        if form.is_valid():
            files = request.FILES.getlist('doc_files')

            results = {
                'success': 0,
                'skipped': 0,
                'error': 0,
                'errors': []
            }

            temp_dir = os.path.join(settings.MEDIA_ROOT, 'temp')
            os.makedirs(temp_dir, exist_ok=True)

            for file in files:
                file_path = os.path.join(temp_dir, file.name)

                try:
                    # Сохраняем файл
                    with open(file_path, 'wb+') as destination:
                        for chunk in file.chunks():
                            destination.write(chunk)

                    # Парсим документ - ВСЯ ЛОГИКА ПАРСИНГА ВНУТРИ try/finally
                    try:
                        # Импортируем парсер
                        from .utils.doc_parser import DocParser

                        # Парсим документ
                        parser = DocParser(file_path)
                        protocol_data = parser.parse()

                        # Явно закрываем документы, если есть такая возможность
                        if hasattr(parser, 'doc'):
                            parser.doc = None
                        if hasattr(parser, 'doc_text'):
                            parser.doc_text = None

                        # Проверяем, есть ли номер протокола
                        if protocol_data.get('number_protocol'):
                            # Проверяем, существует ли протокол в БД
                            exists = Protocol.objects.filter(
                                number_protocol=protocol_data['number_protocol']
                            ).exists()

                            if exists:
                                results['skipped'] += 1
                                messages.warning(
                                    request,
                                    f'⏭ Протокол №{protocol_data["number_protocol"]} уже существует в БД'
                                )
                            else:
                                # Создаем новый протокол
                                protocol = Protocol(**protocol_data)
                                protocol.save()
                                results['success'] += 1
                                messages.success(
                                    request,
                                    f'✅ Протокол №{protocol_data["number_protocol"]} успешно добавлен'
                                )
                        else:
                            results['error'] += 1
                            results['errors'].append(f"{file.name}: Не удалось извлечь номер протокола")
                            messages.error(
                                request,
                                f'❌ {file.name}: Не удалось извлечь номер протокола'
                            )

                    except ImportError as e:
                        results['error'] += 1
                        error_msg = f"{file.name}: Ошибка импорта - {str(e)}. Установите python-docx и docx2python"
                        results['errors'].append(error_msg)
                        messages.error(request, f'❌ {error_msg}')

                    except Exception as e:
                        results['error'] += 1
                        results['errors'].append(f"{file.name}: {str(e)}")
                        messages.error(request, f'❌ Ошибка при обработке {file.name}: {str(e)}')

                    finally:
                        # Принудительно вызываем сборщик мусора
                        import gc
                        gc.collect()

                        # Удаляем временный файл
                        if os.path.exists(file_path):
                            try:
                                # Даем время на освобождение файла
                                import time
                                time.sleep(0.1)
                                os.remove(file_path)
                            except PermissionError as e:
                                # Если файл все еще занят, попробуем переименовать его и удалить позже
                                try:
                                    import uuid
                                    temp_rename = os.path.join(temp_dir, f"to_delete_{uuid.uuid4()}.tmp")
                                    os.rename(file_path, temp_rename)
                                    os.remove(temp_rename)
                                except:
                                    # Если не получается, запланируем удаление при следующем запуске
                                    results['errors'].append(f"{file.name}: Не удалось удалить временный файл")
                            except Exception as e:
                                results['errors'].append(f"{file.name}: Ошибка при удалении: {str(e)}")
                except Exception as e:
                    results['error'] += 1
                    results['errors'].append(f"{file.name}: Ошибка при сохранении: {str(e)}")
                    messages.error(request, f'❌ Ошибка при сохранении {file.name}: {str(e)}')

            # Итоговое сообщение
            summary = (
                f"📊 ИТОГО: Успешно: {results['success']} | "
                f"Пропущено: {results['skipped']} | "
                f"Ошибки: {results['error']}"
            )
            messages.success(request, summary)

            return redirect('protocol_app:upload_protocols')
    else:
        form = ProtocolUploadForm()

    return render(request, 'protocol_app/upload.html', {'form': form})

@require_http_methods(['GET'])
def search_protocols(request):
    """Поиск протоколов"""
    form = ProtocolSearchForm(request.GET or None)
    protocols = Protocol.objects.none()

    context = {
        'form': form,
        'page_obj': [],
        'total_count': 0,
    }
    return render(request, 'protocol_app/search.html', context)


@require_http_methods(['GET'])
def protocol_detail(request, pk):
    """Детальная информация о протоколе"""
    return render(request, 'protocol_app/detail.html', {'protocol': None})


@require_http_methods(['GET'])
def export_single_protocol(request, pk):
    """Экспорт одного протокола в Excel"""
    return HttpResponse(f"Экспорт протокола {pk}")


@require_http_methods(['POST'])
def delete_protocol(request, pk):
    """Удаление протокола"""
    return redirect('protocol_app:search_protocols')


@require_http_methods(['GET', 'POST'])
def export_protocols(request):
    """Экспорт протоколов в Excel/CSV"""
    if request.method == 'POST':
        return HttpResponse("Экспорт всех протоколов")
    else:
        form = ProtocolExportForm()

    context = {
        'form': form,
        'total_count': 0
    }
    return render(request, 'protocol_app/export.html', context)


@require_http_methods(['GET'])
def ajax_search_protocols(request):
    """AJAX поиск протоколов для автокомплита"""
    return JsonResponse([], safe=False)


def export_page(request):
    """Страница экспорта данных"""
    from .models import Protocol  # Импортируем вашу модель

    total_count = Protocol.objects.count()  # Получаем общее количество записей

    # Создаем форму для экспорта
    from django import forms
    class ExportForm(forms.Form):
        EXPORT_FORMATS = [
            ('excel', 'Microsoft Excel (.xlsx)'),
        ]
        export_format = forms.ChoiceField(
            choices=EXPORT_FORMATS,
            widget=forms.RadioSelect,
            initial='excel',
            label='Формат экспорта'
        )
        include_all = forms.BooleanField(
            required=False,
            initial=True,
            label='Экспортировать все протоколы'
        )
        number_protocol = forms.CharField(
            required=False,
            widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Например: 2341-В/25'}),
            label='Номер протокола'
        )

    form = ExportForm()

    return render(request, 'export.html', {
        'form': form,
        'total_count': total_count
    })


@csrf_exempt
def export_to_excel(request):
    """API для экспорта данных в Excel"""

    if request.method != 'POST':
        return HttpResponse("Метод не разрешен", status=405)

    try:
        # Параметры экспорта
        export_all = request.POST.get('include_all') == 'on'
        protocol_number = request.POST.get('number_protocol', '').strip()

        # Путь к базе данных
        db_path = os.path.join(settings.BASE_DIR, 'db.sqlite3')
        engine = create_engine(f'sqlite:///{db_path}')

        # Формируем запрос в зависимости от параметров
        if export_all:
            query = "SELECT * FROM protocols"
            log_msg = "Подготовка выгрузки всех протоколов в Excel..."
        else:
            query = f"SELECT * FROM protocols WHERE number = '{protocol_number}'"
            log_msg = f"Подготовка выгрузки протокола №{protocol_number} в Excel..."

        print(log_msg)  # Для отладки

        # Получаем данные из БД
        df = pd.read_sql_query(query, con=engine)

        # Проверяем наличие файла с маппингом колонок
        columns_file = os.path.join(settings.BASE_DIR, 'Columns_02_top.xlsx')
        if os.path.exists(columns_file):
            # Переименовываем колонки в кириллицу
            df_col = pd.read_excel(columns_file)
            kirill = df_col.old_name.to_list()
            latin = df_col.new_pokazat_name.to_list()

            dict_trans = {}
            for i in latin:
                dict_trans[i] = kirill[latin.index(i)]

            for i in df.columns.to_list():
                if i in dict_trans:
                    df.rename(columns={i: dict_trans[i]}, inplace=True)

        if df.empty:
            if export_all:
                return HttpResponse("База данных пуста!", status=400)
            else:
                return HttpResponse(f"Протокол №{protocol_number} не найден в базе данных!", status=404)

        # Создаем HTTP ответ с Excel файлом
        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        # Формируем имя файла
        if export_all:
            filename = f'protocols_export_all_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        else:
            filename = f'protocol_{protocol_number}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        # Создаем Excel файл с xlsxwriter
        workbook = xlsxwriter.Workbook(response, {'in_memory': True})
        worksheet = workbook.add_worksheet('Протоколы')

        # Создаем форматы
        header_format_a_j = workbook.add_format({
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'bold': True,
            'bg_color': '#44c14d',
            'border': 1,
            'font_size': 9
        })

        header_format_k_hd = workbook.add_format({
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'bold': True,
            'bg_color': '#D9D9D9',
            'border': 1,
            'font_size': 9
        })

        data_format = workbook.add_format({
            'border': 1,
            'font_size': 9
        })

        date_format = workbook.add_format({
            'border': 1,
            'font_size': 9,
            'num_format': 'DD.MM.YYYY'
        })

        # Получаем заголовки
        headers = list(df.columns)
        num_cols = len(headers)
        num_rows = len(df)

        # Записываем заголовки
        for col_num, header in enumerate(headers):
            if col_num <= 9:
                worksheet.write(0, col_num, header, header_format_a_j)
            else:
                worksheet.write(0, col_num, header, header_format_k_hd)

        # Записываем данные
        for row_num in range(num_rows):
            for col_num in range(num_cols):
                cell_value = df.iat[row_num, col_num]

                if isinstance(cell_value, (pd.Timestamp, datetime)):
                    worksheet.write(row_num + 1, col_num, cell_value, date_format)
                else:
                    worksheet.write(row_num + 1, col_num, cell_value, data_format)

        # Устанавливаем высоту строки для заголовка
        worksheet.set_row(0, 30)

        # Устанавливаем ширину колонок A-J
        column_widths = {
            0: 15, 1: 15, 2: 15, 3: 30, 4: 30,
            5: 30, 6: 12, 7: 18, 8: 20, 9: 15,
        }

        for col_num, width in column_widths.items():
            worksheet.set_column(col_num, col_num, width)

        # Автоматическая ширина для остальных колонок
        for col_num in range(num_cols):
            if col_num not in column_widths:
                max_length = 0
                header_len = len(str(headers[col_num]))
                max_length = max(max_length, header_len)

                for row_num in range(num_rows):
                    cell_value = df.iat[row_num, col_num]
                    if cell_value is not None:
                        cell_len = len(str(cell_value))
                        max_length = max(max_length, cell_len)

                worksheet.set_column(col_num, col_num, min(max_length + 2, 50))

        # Устанавливаем фильтр
        last_col_letter = xlsxwriter.utility.xl_col_to_name(num_cols - 1)
        filter_range = f'A1:{last_col_letter}{num_rows + 1}'
        worksheet.autofilter(filter_range)

        # Закрепляем заголовок
        worksheet.freeze_panes(1, 0)

        workbook.close()

        return response

    except Exception as e:
        logger.error(f"Ошибка при выгрузке в Excel: {str(e)}")
        return HttpResponse(f"Ошибка при выгрузке: {str(e)}", status=500)

