from django.db import models


class Protocol(models.Model):
    # Основная информация
    rejd_na_istochnik = models.CharField('Рейд на источник', max_length=100, blank=True, null=True, default='-')
    number_protocol = models.CharField('Номер протокола', max_length=100, db_index=True)
    date_protocol = models.CharField('Дата протокола', max_length=50, blank=True, null=True)

    # Адресная информация
    okrug = models.CharField('Округ', max_length=100, blank=True, null=True)
    region = models.CharField('Район', max_length=100, blank=True, null=True)
    nazvanie_territorii = models.CharField('Название территории', max_length=300, blank=True, null=True)

    # Дата и время измерений
    data_izmereniya = models.CharField('Дата измерения', max_length=50, blank=True, null=True)
    time_start_izmereniya = models.CharField('Время начала', max_length=50, blank=True, null=True)
    time_end_izmereniya = models.CharField('Время завершения', max_length=50, blank=True, null=True)

    # Тип протокола
    type_protocol = models.CharField('Тип протокола', max_length=50, blank=True, null=True)

    # Газы и вещества
    dioksid_sery = models.CharField('Диоксид серы', max_length=50, blank=True, null=True)
    serovodorod = models.CharField('Сероводород', max_length=50, blank=True, null=True)
    oksid_azota = models.CharField('Оксид азота', max_length=50, blank=True, null=True)
    dioksid_azota = models.CharField('Диоксид азота', max_length=50, blank=True, null=True)
    oksid_ugleroda = models.CharField('Оксид углерода', max_length=50, blank=True, null=True)
    metan = models.CharField('Метан', max_length=50, blank=True, null=True)
    sum_uglevodorodov_v_pereschete_na_metan = models.CharField('Сумма углеводородов в пересчете на метан',
                                                               max_length=50, blank=True, null=True)
    sum_uglevodorodov_za_vychetom_metana = models.CharField('Сумма углеводородов за вычетом метана', max_length=50,
                                                            blank=True, null=True)
    vzveshennye_chasticy_10_mkm = models.CharField('Взвешенные частицы 10 мкм', max_length=50, blank=True, null=True)
    vzveshennye_chasticy_25_mkm = models.CharField('Взвешенные частицы 2.5 мкм', max_length=50, blank=True, null=True)
    ammiak_pel = models.CharField('Аммиак (ПЭЛ)', max_length=50, blank=True, null=True)
    ammiak_ai = models.CharField('Аммиак (АИ)', max_length=50, blank=True, null=True)
    pyl_70sio220 = models.CharField('Пыль 70% SiO2 20-30%', max_length=50, blank=True, null=True)
    vzveshennye_veshchestva = models.CharField('Взвешенные вещества', max_length=50, blank=True, null=True)
    ftorid_vodoroda = models.CharField('Фторид водорода', max_length=50, blank=True, null=True)
    hlorid_vodoroda = models.CharField('Хлорид водорода', max_length=50, blank=True, null=True)
    hlor = models.CharField('Хлор', max_length=50, blank=True, null=True)
    benzapiren_34benzpiren = models.CharField('Бензапирен (3,4-Бензпирен)', max_length=50, blank=True, null=True)
    fenol = models.CharField('Фенол', max_length=50, blank=True, null=True)
    formaldegid = models.CharField('Формальдегид', max_length=50, blank=True, null=True)
    benzin = models.CharField('Бензин', max_length=50, blank=True, null=True)
    akrilonitril = models.CharField('Акрилонитрил', max_length=50, blank=True, null=True)
    anilin = models.CharField('Анилин', max_length=50, blank=True, null=True)
    aceton = models.CharField('Ацетон', max_length=50, blank=True, null=True)
    benzol = models.CharField('Бензол', max_length=50, blank=True, null=True)
    butanol = models.CharField('Бутанол', max_length=50, blank=True, null=True)
    butilacetat = models.CharField('Бутилацетат', max_length=50, blank=True, null=True)
    vinilacetat = models.CharField('Винилацетат', max_length=50, blank=True, null=True)
    geksan = models.CharField('Гексан', max_length=50, blank=True, null=True)
    geptan = models.CharField('Гептан', max_length=50, blank=True, null=True)
    dekan = models.CharField('Декан', max_length=50, blank=True, null=True)
    dihlormetan = models.CharField('Дихлорметан', max_length=50, blank=True, null=True)
    izoamilacetat = models.CharField('Изоамилацетат', max_length=50, blank=True, null=True)
    izobutanol = models.CharField('Изобутанол', max_length=50, blank=True, null=True)
    izopropanol = models.CharField('Изопропанол', max_length=50, blank=True, null=True)
    m_pksiloly = models.CharField('м-П-Ксилолы', max_length=50, blank=True, null=True)
    oksilol = models.CharField('Оксилол', max_length=50, blank=True, null=True)
    izopropilbenzol_kumol = models.CharField('Изопропилбензол (Кумол)', max_length=50, blank=True, null=True)
    mezitilen = models.CharField('Мезетилен', max_length=50, blank=True, null=True)
    metanol = models.CharField('Метанол', max_length=50, blank=True, null=True)
    metilstirol = models.CharField('Метилстирол', max_length=50, blank=True, null=True)
    metiletilketon = models.CharField('Метилэтилкетон', max_length=50, blank=True, null=True)
    nonan = models.CharField('Нонан', max_length=50, blank=True, null=True)
    oktan = models.CharField('Октан', max_length=50, blank=True, null=True)
    propanol = models.CharField('Пропанол', max_length=50, blank=True, null=True)
    propilbenzol = models.CharField('Пропилбензол', max_length=50, blank=True, null=True)
    spirt_amilovyj = models.CharField('Спирт амиловый', max_length=50, blank=True, null=True)
    spirt_izoamilovyj = models.CharField('Спирт изоамиловый', max_length=50, blank=True, null=True)
    stirol = models.CharField('Стирол', max_length=50, blank=True, null=True)
    tetrahloretilen = models.CharField('Тетрахлорэтилен', max_length=50, blank=True, null=True)
    toluol = models.CharField('Толуол', max_length=50, blank=True, null=True)
    trihloretilen = models.CharField('Трихлорэтилен', max_length=50, blank=True, null=True)
    uglerod_4xhloristyj = models.CharField('Углерод 4-хлористый', max_length=50, blank=True, null=True)
    hlorbenzol = models.CharField('Хлорбензол', max_length=50, blank=True, null=True)
    ciklogeksanon = models.CharField('Циклогексанон', max_length=50, blank=True, null=True)
    etanol = models.CharField('Этанол', max_length=50, blank=True, null=True)
    etilacetat = models.CharField('Этилацетат', max_length=50, blank=True, null=True)
    etilcellozolv = models.CharField('Этилцеллозольв', max_length=50, blank=True, null=True)
    metilmerkaptan = models.CharField('Метилмеркаптан', max_length=50, blank=True, null=True)
    etilmerkaptan = models.CharField('Этилмеркаптан', max_length=50, blank=True, null=True)
    propilmerkaptan = models.CharField('Пропилмеркаптан', max_length=50, blank=True, null=True)
    butilmerkaptan = models.CharField('Бутилмеркаптан', max_length=50, blank=True, null=True)
    izopropilmerkaptan = models.CharField('Изопропилмеркаптан', max_length=50, blank=True, null=True)
    vtorbutilmerkaptan = models.CharField('Втор-бутилмеркаптан', max_length=50, blank=True, null=True)
    tretbutilmerkaptan = models.CharField('Трет-бутилмеркаптан', max_length=50, blank=True, null=True)
    izobutilmerkaptan = models.CharField('Изобутилмеркаптан', max_length=50, blank=True, null=True)

    # Металлы
    alyuminij = models.CharField('Алюминий', max_length=50, blank=True, null=True)
    barij = models.CharField('Барий', max_length=50, blank=True, null=True)
    berillij = models.CharField('Бериллий', max_length=50, blank=True, null=True)
    vanadij = models.CharField('Ванадий', max_length=50, blank=True, null=True)
    zhelezo = models.CharField('Железо', max_length=50, blank=True, null=True)
    kadmij = models.CharField('Кадмий', max_length=50, blank=True, null=True)
    kobalt = models.CharField('Кобальт', max_length=50, blank=True, null=True)
    kremnij = models.CharField('Кремний', max_length=50, blank=True, null=True)
    magnij = models.CharField('Магний', max_length=50, blank=True, null=True)
    marganec = models.CharField('Марганец', max_length=50, blank=True, null=True)
    med = models.CharField('Медь', max_length=50, blank=True, null=True)
    molibden = models.CharField('Молибден', max_length=50, blank=True, null=True)
    myshyak = models.CharField('Мышьяк', max_length=50, blank=True, null=True)
    nikel = models.CharField('Никель', max_length=50, blank=True, null=True)
    olovo = models.CharField('Олово', max_length=50, blank=True, null=True)
    svinec = models.CharField('Свинец', max_length=50, blank=True, null=True)
    serebro = models.CharField('Серебро', max_length=50, blank=True, null=True)
    selen = models.CharField('Селен', max_length=50, blank=True, null=True)
    surma = models.CharField('Сурьма', max_length=50, blank=True, null=True)
    titan = models.CharField('Титан', max_length=50, blank=True, null=True)
    hrom = models.CharField('Хром', max_length=50, blank=True, null=True)
    cink = models.CharField('Цинк', max_length=50, blank=True, null=True)

    # Продолжение списка веществ (сократил для читаемости, нужно добавить все поля)
    pyl_10sio22 = models.CharField('Пыль 10% SiO2', max_length=50, blank=True, null=True)
    pyl_20sio210 = models.CharField('Пыль 20% SiO2 10-20%', max_length=50, blank=True, null=True)
    pyl_sio22 = models.CharField('Пыль SiO2 2-10%', max_length=50, blank=True, null=True)
    dioksid_sery_rd = models.CharField('Диоксид серы РД', max_length=50, blank=True, null=True)
    ugleroda_monooksid_rd = models.CharField('Углерода монооксид РД', max_length=50, blank=True, null=True)
    sazha = models.CharField('Сажа', max_length=50, blank=True, null=True)
    sernaya_kislota_i_sulfaty = models.CharField('Серная кислота и сульфаты', max_length=50, blank=True, null=True)
    karbonilsulfid = models.CharField('Карбонилсульфид', max_length=50, blank=True, null=True)
    disulfid_ugleroda = models.CharField('Дисульфид углерода', max_length=50, blank=True, null=True)
    dimetilsulfid = models.CharField('Диметилсульфид', max_length=50, blank=True, null=True)
    dimetildisulfid = models.CharField('Диметилдисульфид', max_length=50, blank=True, null=True)
    tiofen = models.CharField('Тиофен', max_length=50, blank=True, null=True)
    benzaldegid = models.CharField('Бензальдегид', max_length=50, blank=True, null=True)
    etan = models.CharField('Этан', max_length=50, blank=True, null=True)
    propan = models.CharField('Пропан', max_length=50, blank=True, null=True)
    butan = models.CharField('Бутан', max_length=50, blank=True, null=True)
    pentan = models.CharField('Пентан', max_length=50, blank=True, null=True)
    eten = models.CharField('Этен', max_length=50, blank=True, null=True)
    propen = models.CharField('Пропен', max_length=50, blank=True, null=True)
    izobutan = models.CharField('Изобутан', max_length=50, blank=True, null=True)
    izobuten = models.CharField('Изобутен', max_length=50, blank=True, null=True)
    izopentan = models.CharField('Изопентан', max_length=50, blank=True, null=True)
    npentan = models.CharField('н-Пентан', max_length=50, blank=True, null=True)
    naftalin = models.CharField('Нафталин', max_length=50, blank=True, null=True)
    metilciklogeksan = models.CharField('Метилциклогексан', max_length=50, blank=True, null=True)
    etilenglikol = models.CharField('Этиленгликоль', max_length=50, blank=True, null=True)
    uksusnaya_kislota = models.CharField('Уксусная кислота', max_length=50, blank=True, null=True)

    # Служебные поля
    created_at = models.DateTimeField('Дата создания', auto_now_add=True)
    updated_at = models.DateTimeField('Дата обновления', auto_now=True)

    class Meta:
        db_table = 'protocols'
        verbose_name = 'Протокол'
        verbose_name_plural = 'Протоколы'
        indexes = [
            models.Index(fields=['number_protocol']),
            models.Index(fields=['date_protocol']),
            models.Index(fields=['type_protocol']),
            models.Index(fields=['okrug']),
            models.Index(fields=['region']),
        ]

    def __str__(self):
        return f"{self.number_protocol} - {self.date_protocol}"