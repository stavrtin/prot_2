from django import forms
from django.forms.widgets import FileInput
from .models import Protocol


class MultipleFileInput(FileInput):
    """Кастомный виджет для загрузки нескольких файлов"""
    allow_multiple_selected = True

    def __init__(self, attrs=None):
        if attrs is None:
            attrs = {}
        attrs['multiple'] = 'multiple'
        super().__init__(attrs)

    def value_from_datadict(self, data, files, name):
        if hasattr(files, 'getlist'):
            return files.getlist(name)
        return [super().value_from_datadict(data, files, name)]


class MultipleFileField(forms.FileField):
    """Кастомное поле для загрузки нескольких файлов"""
    widget = MultipleFileInput

    def clean(self, data, initial=None):
        if not data and self.required:
            raise forms.ValidationError(self.error_messages['required'])

        if isinstance(data, list):
            single_file_clean = super().clean
            cleaned_data = []
            for file in data:
                cleaned_data.append(single_file_clean(file, initial))
            return cleaned_data
        return [super().clean(data, initial)]


class ProtocolUploadForm(forms.Form):
    """Форма для загрузки файлов протоколов"""
    doc_files = MultipleFileField(
        label='Выберите файлы протоколов',
        required=True,
        widget=MultipleFileInput(attrs={
            'accept': '.docx,.doc',
            'class': 'form-control'
        })
    )


class ProtocolSearchForm(forms.Form):
    """Форма для поиска протоколов"""
    number_protocol = forms.CharField(
        label='Номер протокола',
        required=False,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Введите номер или часть номера...'
        })
    )
    date_from = forms.DateField(
        label='Дата с',
        required=False,
        widget=forms.DateInput(attrs={
            'class': 'form-control',
            'type': 'date'
        })
    )
    date_to = forms.DateField(
        label='Дата по',
        required=False,
        widget=forms.DateInput(attrs={
            'class': 'form-control',
            'type': 'date'
        })
    )
    okrug = forms.CharField(
        label='Округ',
        required=False,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Введите округ...'
        })
    )
    type_protocol = forms.ChoiceField(
        label='Тип протокола',
        required=False,
        choices=[('', 'Все')],
        widget=forms.Select(attrs={'class': 'form-control'})
    )


class ProtocolExportForm(forms.Form):
    """Форма для экспорта протоколов"""
    EXPORT_FORMAT_CHOICES = [
        ('excel', 'Excel (.xlsx)'),
        ('csv', 'CSV (.csv)'),
    ]

    export_format = forms.ChoiceField(
        label='Формат файла',
        choices=EXPORT_FORMAT_CHOICES,
        initial='excel',
        widget=forms.RadioSelect
    )

    include_all = forms.BooleanField(
        label='Экспортировать все записи',
        required=False,
        initial=True
    )

    number_protocol = forms.CharField(
        label='Номер протокола (для одного протокола)',
        required=False,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Введите номер протокола...'
        })
    )