// Основные функции для работы с протоколами

$(document).ready(function() {
    // Автокомплит для поиска протоколов
    $('#id_number_protocol').autocomplete({
        source: function(request, response) {
            $.ajax({
                url: '/ajax/search/',
                data: {
                    q: request.term
                },
                success: function(data) {
                    response($.map(data, function(item) {
                        return {
                            label: item.number_protocol + ' (' + item.date_protocol + ')',
                            value: item.number_protocol
                        };
                    }));
                }
            });
        },
        minLength: 2,
        delay: 300
    });

    // Валидация файлов при загрузке
    $('#id_doc_files').on('change', function() {
        var files = $(this)[0].files;
        var validExtensions = ['docx'];
        var maxSize = 10 * 1024 * 1024; // 10 MB

        for (var i = 0; i < files.length; i++) {
            var file = files[i];
            var extension = file.name.split('.').pop().toLowerCase();

            if (validExtensions.indexOf(extension) === -1) {
                alert('Файл ' + file.name + ' имеет недопустимый формат. Разрешены только .docx');
                $(this).val('');
                return false;
            }

            if (file.size > maxSize) {
                alert('Файл ' + file.name + ' превышает максимальный размер (10 МБ)');
                $(this).val('');
                return false;
            }
        }
    });

    // Подтверждение удаления
    $('.delete-protocol').on('click', function(e) {
        if (!confirm('Вы уверены, что хотите удалить этот протокол?')) {
            e.preventDefault();
        }
    });

    // Форматирование таблицы результатов
    $('.table-hover tbody tr').hover(
        function() { $(this).addClass('bg-light'); },
        function() { $(this).removeClass('bg-light'); }
    );

    // Автоматическое скрытие сообщений через 5 секунд
    setTimeout(function() {
        $('.alert').fadeOut('slow');
    }, 5000);
});

// Функция для копирования текста в буфер обмена
function copyToClipboard(text) {
    var textarea = document.createElement('textarea');
    textarea.textContent = text;
    textarea.style.position = 'fixed';
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    alert('Скопировано в буфер обмена');
}

// Функция для экспорта таблицы в Excel
function exportTableToExcel(tableId, filename = 'export.xlsx') {
    var table = document.getElementById(tableId);
    var wb = XLSX.utils.table_to_book(table, {sheet: "Sheet1"});
    XLSX.writeFile(wb, filename);
}