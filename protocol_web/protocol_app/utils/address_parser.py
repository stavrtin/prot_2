import re


def parse_moscow_address(address_str):
    """
    Разбирает московский адрес на компоненты:
    - Округ
    - Район
    - Название территории
    """
    result = {
        'okrug': '-',
        'region': '-',
        'nazvanie_territorii': '-'
    }

    if not address_str or not isinstance(address_str, str):
        return result

    # Чистим строку
    address_str = ' '.join(address_str.split())
    address_str = address_str.replace('г. Москва,', '').replace('г. Москва', '').strip()
    if address_str.startswith(','):
        address_str = address_str[1:].strip()

    parts = [part.strip() for part in address_str.split(',')]

    # Список округов
    okrug_abbr = {'ВАО', 'ЗАО', 'САО', 'СЗАО', 'СВАО', 'ЮАО', 'ЮВАО', 'ЮЗАО',
                  'ЗелАО', 'ТиНАО', 'НАО', 'ЦАО'}

    # Типы улиц
    street_types = {'ул.', 'ул', 'улица', 'просп.', 'проспект', 'пр.', 'пр-кт',
                    'ш.', 'шоссе', 'пер.', 'переулок', 'б-р', 'бульвар',
                    'наб.', 'набережная', 'аллея', 'пл.', 'площадь', 'проезд'}

    # Поиск округа
    okrug_found = None
    okrug_index = -1

    for i, part in enumerate(parts):
        if part in okrug_abbr:
            result['okrug'] = part
            okrug_index = i
            break

    # Поиск района
    if okrug_index != -1 and okrug_index + 1 < len(parts):
        for i in range(okrug_index + 1, min(okrug_index + 3, len(parts))):
            part = parts[i]

            if 'район' in part.lower() or 'поселение' in part.lower():
                district_name = part
                for word in ['район', 'поселение', 'муниципальный']:
                    district_name = district_name.replace(word, '').replace('  ', ' ').strip()
                result['region'] = district_name
                break
            elif not any(word in part.lower() for word in street_types) and 'д.' not in part.lower():
                if part.endswith(('ский', 'вое', 'во', 'ино', 'ово', 'ево')):
                    result['region'] = part
                    break

    # Поиск территории
    territory_parts = []
    start_idx = okrug_index + 1 if okrug_index != -1 else 0

    if result['region'] != '-':
        for i in range(start_idx, len(parts)):
            if result['region'] in parts[i]:
                start_idx = i + 1
                break

    for i in range(start_idx, len(parts)):
        part = parts[i]
        if any(prefix in part.lower() for prefix in street_types) or \
                (territory_parts and any(word in part.lower() for word in ['д.', 'дом', 'к.', 'стр.'])):
            territory_parts.append(part)
        elif i == start_idx and not any(word in part.lower() for word in ['район', 'поселение']):
            territory_parts.append(part)

    if territory_parts:
        result['nazvanie_territorii'] = ', '.join(territory_parts)

    return result