import os
import pandas as pd
from django.shortcuts import render, redirect
from django.http import HttpResponse, JsonResponse
from django.contrib import messages
from django.conf import settings
from django.views.decorators.http import require_http_methods
from django.core.paginator import Paginator
from django.db.models import Q
import xlsxwriter
from io import BytesIO
from datetime import datetime

from .models import Protocol
from .forms import ProtocolUploadForm, ProtocolSearchForm, ProtocolExportForm
from .utils.doc_parser import DocParser


def index(request):
    """Главная страница"""
    total_protocols = Protocol.objects.count()
    pel_count = Protocol.objects.filter(type_protocol__startswith='pel').count()
    ai_count = Protocol.objects.filter(type_protocol__startswith='ai').count()

    # Последние 10 протоколов
    recent_protocols = Protocol.objects.order_by('-created_at')[:10]

    context = {
        'total_protocols': total_protocols,
        'pel_count': pel_count,
        'ai_count': ai_count,
        'recent_protocols': recent_protocols,
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

                    # Импортируем парсер
                    from .utils.doc_parser import DocParser

                    # Парсим документ
                    parser = DocParser(file_path)
                    protocol_data = parser.parse()

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

                except Exception as e:
                    results['error'] += 1
                    results['errors'].append(f"{file.name}: {str(e)}")
                    messages.error(request, f'❌ Ошибка при обработке {file.name}: {str(e)}')

                finally:
                    # Удаляем временный файл
                    if os.path.exists(file_path):
                        os.remove(file_path)

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


@require_http_methods(['GET', 'POST'])
def search_protocols(request):
    """Поиск протоколов"""
    form = ProtocolSearchForm(request.GET or None)
    protocols = Protocol.objects.all()
    query_params = {}

    if form.is_valid():
        number = form.cleaned_data.get('number_protocol')
        date_from = form.cleaned_data.get('date_from')
        date_to = form.cleaned_data.get('date_to')
        okrug = form.cleaned_data.get('okrug')
        type_protocol = form.cleaned_data.get('type_protocol')

        if number:
            protocols = protocols.filter(number_protocol__icontains=number)
            query_params['number'] = number

        if date_from:
            protocols = protocols.filter(date_protocol__gte=date_from)
            query_params['date_from'] = date_from

        if date_to:
            protocols = protocols.filter(date_protocol__lte=date_to)
            query_params['date_to'] = date_to

        if okrug:
            protocols = protocols.filter(okrug__icontains=okrug)
            query_params['okrug'] = okrug

        if type_protocol:
            protocols = protocols.filter(type_protocol=type_protocol)
            query_params['type'] = type_protocol

    # Пагинация
    paginator = Paginator(protocols.order_by('-date_protocol', '-number_protocol'), 20)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    context = {
        'form': form,
        'page_obj': page_obj,
        'query_params': query_params,
        'total_count': protocols.count(),
    }
    return render(request, 'protocol_app/search.html', context)


@require_http_methods(['GET'])
def protocol_detail(request, pk):
    """Детальная информация о протоколе"""
    protocol = Protocol.objects.get(pk=pk)
    return render(request, 'protocol_app/detail.html', {'protocol': protocol})


@require_http_methods(['GET', 'POST'])
def export_protocols(request):
    """Экспорт протоколов в Excel/CSV"""
    if request.method == 'POST':
        form = ProtocolExportForm(request.POST)
        if form.is_valid():
            export_format = form.cleaned_data['export_format']
            include_all = form.cleaned_data['include_all']
            number_protocol = form.cleaned_data['number_protocol']

            # Получаем данные
            if not include_all and number_protocol:
                protocols = Protocol.objects.filter(number_protocol=number_protocol)
            else:
                protocols = Protocol.objects.all()

            if not protocols.exists():
                messages.warning(request, 'Нет данных для экспорта')
                return redirect('export_protocols')

            # Преобразуем в DataFrame
            data = []
            for p in protocols:
                data.append({field.name: getattr(p, field.name) for field in Protocol._meta.fields})

            df = pd.DataFrame(data)

            # Создаем файл
            if export_format == 'excel':
                output = BytesIO()
                workbook = xlsxwriter.Workbook(output)
                worksheet = workbook.add_worksheet('Протоколы')

                # Форматы
                header_format = workbook.add_format({
                    'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                    'bold': True, 'bg_color': '#44c14d', 'border': 1, 'font_size': 9
                })
                data_format = workbook.add_format({'border': 1, 'font_size': 9})

                # Заголовки
                headers = list(df.columns)
                for col_num, header in enumerate(headers):
                    worksheet.write(0, col_num, header, header_format)

                # Данные
                for row_num in range(len(df)):
                    for col_num in range(len(headers)):
                        worksheet.write(row_num + 1, col_num, df.iat[row_num, col_num], data_format)

                worksheet.set_row(0, 30)
                worksheet.freeze_panes(1, 0)

                workbook.close()

                output.seek(0)

                filename = f'protocols_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
                response = HttpResponse(
                    output.read(),
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                response['Content-Disposition'] = f'attachment; filename="{filename}"'
                return response

            elif export_format == 'csv':
                response = HttpResponse(content_type='text/csv')
                filename = f'protocols_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
                response['Content-Disposition'] = f'attachment; filename="{filename}"'

                df.to_csv(response, index=False, encoding='utf-8-sig')
                return response
    else:
        form = ProtocolExportForm()

    context = {
        'form': form,
        'total_count': Protocol.objects.count()
    }
    return render(request, 'protocol_app/export.html', context)


@require_http_methods(['GET'])
def export_single_protocol(request, pk):
    """Экспорт одного протокола в Excel"""
    try:
        protocol = Protocol.objects.get(pk=pk)

        # Создаем DataFrame
        data = {field.name: [getattr(protocol, field.name)] for field in Protocol._meta.fields}
        df = pd.DataFrame(data)

        # Транспонируем для вертикального отображения
        df_one = df.T.reset_index().rename(columns={'index': 'Параметр', 0: 'Значение'})
        df_one = df_one.loc[~df_one['Значение'].isna()]

        # Создаем Excel
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet('Протокол')

        # Форматы
        title_format = workbook.add_format({
            'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'
        })
        header_format = workbook.add_format({
            'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'bold': True, 'bg_color': '#44c14d', 'border': 1, 'font_size': 9
        })
        data_format = workbook.add_format({'border': 1, 'font_size': 9})

        # Заголовок
        worksheet.merge_range('A1:B1', f'Протокол №{protocol.number_protocol}', title_format)

        # Заголовки колонок
        headers = list(df_one.columns)
        for col_num, hdr in enumerate(headers):
            worksheet.write(1, col_num, hdr, header_format)

        # Данные
        for row_idx in range(len(df_one)):
            for col_idx in range(len(headers)):
                val = df_one.iat[row_idx, col_idx]
                worksheet.write(row_idx + 2, col_idx, val, data_format)

        worksheet.set_row(0, 25)
        worksheet.set_row(1, 30)
        worksheet.freeze_panes(2, 0)

        workbook.close()

        output.seek(0)

        filename = f'protocol_{protocol.number_protocol}.xlsx'
        response = HttpResponse(
            output.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response

    except Protocol.DoesNotExist:
        messages.error(request, 'Протокол не найден')
        return redirect('search_protocols')


@require_http_methods(['POST'])
def delete_protocol(request, pk):
    """Удаление протокола"""
    try:
        protocol = Protocol.objects.get(pk=pk)
        protocol.delete()
        messages.success(request, f'Протокол №{protocol.number_protocol} успешно удален')
    except Protocol.DoesNotExist:
        messages.error(request, 'Протокол не найден')

    return redirect('search_protocols')


@require_http_methods(['GET'])
def ajax_search_protocols(request):
    """AJAX поиск протоколов для автокомплита"""
    query = request.GET.get('q', '')
    protocols = Protocol.objects.filter(number_protocol__icontains=query)[:10]

    results = []
    for p in protocols:
        results.append({
            'id': p.id,
            'number_protocol': p.number_protocol,
            'date_protocol': p.date_protocol,
            'type_protocol': p.type_protocol
        })

    return JsonResponse(results, safe=False)