import os
import sys
import warnings
import json
import copy
import requests
import re

import pandas as pd
from urllib3.exceptions import InsecureRequestWarning

from _config import SCRIPT_DIR, FILES_DIR, EXCEL_FILE_NAME, BASE_URL, LICENSES_COLLECTION, RECORDS_COLLECTION, STANDARD_CODES, UNIT, TEST, EXCEL_LISTS, RECORDS_TEMPLATES
from _logger import setup_logger, setup_success_logger, setup_fail_logger
from _utils import (
    nz,
    split_sc,
    read_excel,
    to_iso_date,
    parse_date_to_birthday_obj,
    format_phone,
    format_multiple_phones,
    read_file_as_base64,
    find_file_in_dir,
    find_document_group_by_mnemonic,
    generate_guid,
    jsonable
)
from _api import (
    api_request,
    upload_file,
    delete_file_from_storage,
    setup_session,
    get_subservices,
    get_unit,
    get_standard_code,
    create_appeal_data,
    create_mainElement_data,
    create_subservice_data,
    create_subject_data,
    create_appeal_with_entities,
    delete_from_collection
)
from _templates import SUBJECT_UL, SUBJECT_IP

def normStr(s):
    if s is None:
        return None
    s = str(s).strip()
    return s if s else None

def split_postal_address(s):
    raw = str(s).strip()
    if not raw:
        return {"postalCode": None, "fullAddress": None}

    # "236006, Калининград ..." -> postalCode=236006, fullAddress="Калининград ..."
    m = re.match(r"^(\d{6})(?:,\s*)?(.*)$", raw)
    if not m:
        return {"postalCode": None, "fullAddress": raw}

    postal_code = m.group(1)
    rest = (m.group(2) or "").strip()

    if not rest:
        rest = None

    return {
        "postalCode": postal_code,
        "fullAddress": rest
    }

def main():
    warnings.filterwarnings("ignore", category=InsecureRequestWarning)

    logger = setup_logger()
    successLogger = setup_success_logger()
    failLogger = setup_fail_logger()

    try:
        if not TEST:
            session = setup_session(logger)
            if session is None:
                sys.exit(1)

        excel_path = os.path.join(SCRIPT_DIR, EXCEL_FILE_NAME)
        logger.info(f"Чтение файла: {excel_path}")
        for list_name in EXCEL_LISTS:
            logger.info(f"Обработка листа: {list_name}")

            excel = read_excel(excel_path, skiprows=3, sheet_name=list_name)
            if excel is None:
                logger.error(f"Файл {excel_path} не найден")
                sys.exit(1)
            excel = excel.iloc[1:].reset_index(drop=True) # Удаляем строку с подписями и сбрасываем индекс

            logger.info(f"Загружено строк: {len(excel)}")
            excel.columns = [c.strip() for c in excel.columns]
            rows_total = len(excel)
            for i, row in enumerate(excel.to_dict("records"), start=1):
                logger.info(f"{i}/{rows_total}")

                unit = UNIT

                if TEST:
                    print(row)
                MAP_PERMISSION_STATUS = {
                    "действует":      { "code": "Working",     "name": "Действует" },
                    "приостановлено": { "code": "Stop",        "name": "Приостановлено" },
                    "аннулировано":   { "code": "Annul",       "name": "Аннулировано" },
                    "не действует":   { "code": "doesNotWork", "name": "Не действует" },
                    "черновик":       { "code": "Draft",       "name": "Черновик" }
                }
                MAP_MARKET_TYPE = {
                    "специализированный": { "code": "Specialized", "_id": "68ed17bb27eea1af1d5473f7", "name": "Специализированный" },
                    "универсальный":      { "code": "Universal",   "_id": "68ed1979b12469d98995f24e", "name": "Универсальный" }
                }
                MAP_SPECIALIZATION = {
                    "сельскохозяйственный":                       { "code":"Agricultural", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Сельскохозяйственный" },
                    "сельскохозяйственный кооперативный":        { "code":"AgriculturalCooperative", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Сельскохозяйственный кооперативный" },
                    "вещевой":                                   { "code":"Clothing", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Вещевой" },
                    "по продаже радио- и электробытовой техники": { "code":"For the sale of radio and electrical appliances", "parentId":"68ed17bb27eea1af1d5473f7", "name":"По продаже радио- и электробытовой техники" },
                    "иная":                                      { "code":"Other", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Иная" },
                    "иное":                                      { "code":"Other", "parentId":"68ed17bb27eea1af1d5473f7", "name":"Иная" },
                    "по продаже строительных материалов":         { "code":"SaleBuildingMaterials", "parentId":"68ed17bb27eea1af1d5473f7", "name":"По продаже строительных материалов" },
                    "по продаже продуктов питания":               { "code":"SaleProducts", "parentId":"68ed17bb27eea1af1d5473f7", "name":"По продаже продуктов питания" }
                }
                MAP_OPERATION_PERIOD = {
                    "постоянный": { "code":"PermanentType", "name":"Постоянный" },
                    "временный":  { "code":"TemporaryType", "name":"Временный" }
                }
                MAP_ORG_STATE_FORM = {
                    "негосударственные пенсионные фонды": { "code": "70402", "name": "Негосударственные пенсионные фонды" },
                    "районные суды, городские суды, межрайонные суды (районные суды)": { "code": "30008", "name": "Районные суды, городские суды, межрайонные суды (районные суды)" },
                    "товарищества собственников недвижимости": { "code": "20700", "name": "Товарищества собственников недвижимости" },
                    "обособленные подразделения юридических лиц": { "code": "30003", "name": "Обособленные подразделения юридических лиц" },
                    "адвокаты, учредившие адвокатский кабинет": { "code": "50201", "name": "Адвокаты, учредившие адвокатский кабинет" },
                    "сельскохозяйственные потребительские животноводческие кооперативы": { "code": "20115", "name": "Сельскохозяйственные потребительские животноводческие кооперативы" },
                    "садоводческие или огороднические некоммерческие товарищества": { "code": "20702", "name": "Садоводческие или огороднические некоммерческие товарищества" },
                    "жилищные или жилищно-строительные кооперативы": { "code": "20102", "name": "Жилищные или жилищно-строительные кооперативы" },
                    "казачьи общества, внесенные в государственный реестр казачьих обществ в российской федерации": { "code": "21100", "name": "Казачьи общества, внесенные в государственный реестр казачьих обществ в Российской Федерации" },
                    "сельскохозяйственные производственные кооперативы": { "code": "14100", "name": "Сельскохозяйственные производственные кооперативы" },
                    "нотариусы, занимающиеся частной практикой": { "code": "50202", "name": "Нотариусы, занимающиеся частной практикой" },
                    "крестьянские (фермерские) хозяйства": { "code": "15300", "name": "Крестьянские (фермерские) хозяйства" },
                    "благотворительные учреждения": { "code": "75502", "name": "Благотворительные учреждения" },
                    "прочие юридические лица, являющиеся коммерческими организациями": { "code": "19000", "name": "Прочие юридические лица, являющиеся коммерческими организациями" },
                    "объединения фермерских хозяйств": { "code": "20613", "name": "Объединения фермерских хозяйств" },
                    "нотариальные палаты": { "code": "20610", "name": "Нотариальные палаты" },
                    "некоммерческие партнерства": { "code": "20614", "name": "Некоммерческие партнерства" },
                    "общественные фонды": { "code": "70403", "name": "Общественные фонды" },
                    "федеральные государственные автономные учреждения": { "code": "75101", "name": "Федеральные государственные автономные учреждения" },
                    "сельскохозяйственные потребительские обслуживающие кооперативы": { "code": "20111", "name": "Сельскохозяйственные потребительские обслуживающие кооперативы" },
                    "общественные организации": { "code": "20200", "name": "Общественные  организации" },
                    "государственные корпорации": { "code": "71601", "name": "Государственные корпорации" },
                    "главы крестьянских (фермерских) хозяйств": { "code": "50101", "name": "Главы крестьянских (фермерских) хозяйств" },
                    "автономные некоммерческие организации": { "code": "71400", "name": "Автономные некоммерческие организации" },
                    "учреждения": { "code": "75000", "name": "Учреждения" },
                    "государственные автономные учреждения субъектов российской федерации": { "code": "75201", "name": "Государственные автономные учреждения субъектов Российской Федерации" },
                    "объединения (ассоциации и союзы) благотворительных организаций": { "code": "20620", "name": "Объединения (ассоциации и союзы) благотворительных организаций" },
                    "публичные акционерные общества": { "code": "12247", "name": "Публичные акционерные общества" },
                    "государственные академии наук": { "code": "75300", "name": "Государственные академии наук" },
                    "государственные компании": { "code": "71602", "name": "Государственные компании" },
                    "представительства юридических лиц": { "code": "30001", "name": "Представительства юридических лиц" },
                    "государственные бюджетные учреждения субъектов российской федерации": { "code": "75203", "name": "Государственные бюджетные учреждения субъектов Российской Федерации" },
                    "союзы (ассоциации) кредитных кооперативов": { "code": "20604", "name": "Союзы (ассоциации) кредитных кооперативов" },
                    "межправительственные международные организации": { "code": "40001", "name": "Межправительственные международные организации" },
                    "муниципальные бюджетные учреждения": { "code": "75403", "name": "Муниципальные бюджетные учреждения" },
                    "государственные унитарные предприятия субъектов российской федерации": { "code": "65242", "name": "Государственные унитарные предприятия субъектов Российской Федерации" },
                    "полные товарищества": { "code": "11051", "name": "Полные товарищества" },
                    "общественные движения": { "code": "20210", "name": "Общественные движения" },
                    "рыболовецкие артели (колхозы)": { "code": "14154", "name": "Рыболовецкие артели (колхозы)" },
                    "потребительские общества": { "code": "20107", "name": "Потребительские общества" },
                    "союзы (ассоциации) общин малочисленных народов": { "code": "20607", "name": "Союзы (ассоциации) общин малочисленных народов" },
                    "сельскохозяйственные потребительские перерабатывающие кооперативы": { "code": "20109", "name": "Сельскохозяйственные потребительские перерабатывающие  кооперативы" },
                    "учреждения, созданные российской федерацией": { "code": "75100", "name": "Учреждения, созданные Российской Федерацией" },
                    "производственные кооперативы (кроме сельскохозяйственных производственных кооперативов)": { "code": "14200", "name": "Производственные кооперативы (кроме сельскохозяйственных производственных кооперативов)" },
                    "учреждения, созданные субъектом российской федерации": { "code": "75200", "name": "Учреждения, созданные субъектом Российской Федерации" },
                    "государственные казенные учреждения субъектов российской федерации": { "code": "75204", "name": "Государственные казенные учреждения субъектов Российской Федерации" },
                    "федеральные государственные унитарные предприятия": { "code": "65241", "name": "Федеральные государственные унитарные предприятия" },
                    "саморегулируемые организации": { "code": "20619", "name": "Саморегулируемые организации" },
                    "территориальные общественные самоуправления": { "code": "20217", "name": "Территориальные общественные самоуправления" },
                    "акционерные общества": { "code": "12200", "name": "Акционерные общества" },
                    "кредитные потребительские кооперативы граждан": { "code": "20105", "name": "Кредитные потребительские  кооперативы граждан" },
                    "казенные предприятия субъектов российской федерации": { "code": "65142", "name": "Казенные предприятия субъектов Российской Федерации" },
                    "советы муниципальных образований субъектов российской федерации": { "code": "20603", "name": "Советы муниципальных образований субъектов Российской Федерации" },
                    "сельскохозяйственные потребительские снабженческие кооперативы": { "code": "20112", "name": "Сельскохозяйственные потребительские снабженческие кооперативы" },
                    "ассоциации (союзы)": { "code": "20600", "name": "Ассоциации (союзы)" },
                    "филиалы юридических лиц": { "code": "30002", "name": "Филиалы юридических лиц" },
                    "муниципальные казенные предприятия": { "code": "65143", "name": "Муниципальные казенные предприятия" },
                    "жилищные накопительные кооперативы": { "code": "20103", "name": "Жилищные накопительные кооперативы" },
                    "органы общественной самодеятельности": { "code": "20211", "name": "Органы общественной самодеятельности" },
                    "религиозные организации": { "code": "71500", "name": "Религиозные организации" },
                    "благотворительные фонды": { "code": "70401", "name": "Благотворительные фонды" },
                    "федеральные государственные казенные учреждения": { "code": "75104", "name": "Федеральные государственные казенные учреждения" },
                    "учреждения, созданные муниципальным образованием (муниципальные учреждения)": { "code": "75400", "name": "Учреждения, созданные муниципальным образованием (муниципальные учреждения)" },
                    "общественные учреждения": { "code": "75505", "name": "Общественные учреждения" },
                    "производственные кооперативы (артели)": { "code": "14000", "name": "Производственные кооперативы (артели)" },
                    "муниципальные автономные учреждения": { "code": "75401", "name": "Муниципальные автономные учреждения" },
                    "хозяйственные общества": { "code": "12000", "name": "Хозяйственные общества" },
                    "адвокатские палаты": { "code": "20609", "name": "Адвокатские палаты" },
                    "общества взаимного страхования": { "code": "20108", "name": "Общества взаимного страхования" },
                    "союзы (ассоциации) общественных объединений": { "code": "20606", "name": "Союзы (ассоциации) общественных объединений" },
                    "общества с ограниченной ответственностью": { "code": "12300", "name": "Общества с ограниченной ответственностью" },
                    "хозяйственные партнерства": { "code": "13000", "name": "Хозяйственные партнерства" },
                    "структурные подразделения обособленных подразделений юридических лиц": { "code": "30004", "name": "Структурные подразделения обособленных подразделений юридических лиц" },
                    "простые товарищества": { "code": "30006", "name": "Простые товарищества" },
                    "коллегии адвокатов": { "code": "20616", "name": "Коллегии адвокатов" },
                    "торгово-промышленные палаты": { "code": "20611", "name": "Торгово-промышленные палаты" },
                    "индивидуальные предприниматели": { "code": "50102", "name": "Индивидуальные предприниматели" },
                    "отделения иностранных некоммерческих неправительственных организаций": { "code": "71610", "name": "Отделения иностранных некоммерческих неправительственных организаций" },
                    "гаражные и гаражно-строительные кооперативы": { "code": "20101", "name": "Гаражные и гаражно-строительные кооперативы" },
                    "частные учреждения": { "code": "75500", "name": "Частные учреждения" },
                    "экологические фонды": { "code": "70404", "name": "Экологические фонды" },
                    "неправительственные международные организации": { "code": "40002", "name": "Неправительственные международные организации" },
                    "союзы потребительских обществ": { "code": "20608", "name": "Союзы потребительских обществ" },
                    "федеральные казенные предприятия": { "code": "65141", "name": "Федеральные казенные предприятия" },
                    "потребительские кооперативы": { "code": "20100", "name": "Потребительские кооперативы" },
                    "фонды проката": { "code": "20121", "name": "Фонды проката" },
                    "публично-правовые компании": { "code": "71600", "name": "Публично-правовые компании" },
                    "фонды": { "code": "70400", "name": "Фонды" },
                    "федеральные государственные бюджетные учреждения": { "code": "75103", "name": "Федеральные государственные бюджетные учреждения" },
                    "товарищества на вере (коммандитные товарищества)": { "code": "11064", "name": "Товарищества на вере (коммандитные товарищества)" },
                    "кредитные потребительские кооперативы": { "code": "20104", "name": "Кредитные потребительские кооперативы" },
                    "товарищества собственников жилья": { "code": "20716", "name": "Товарищества собственников жилья" },
                    "общины коренных малочисленных народов российской федерации": { "code": "21200", "name": "Общины коренных малочисленных народов Российской Федерации" },
                    "хозяйственные товарищества": { "code": "11000", "name": "Хозяйственные товарищества" },
                    "паевые инвестиционные фонды": { "code": "30005", "name": "Паевые инвестиционные фонды" },
                    "политические партии": { "code": "20201", "name": "Политические партии" },
                    "объединения работодателей": { "code": "20612", "name": "Объединения работодателей" },
                    "сельскохозяйственные потребительские сбытовые (торговые) кооперативы": { "code": "20110", "name": "Сельскохозяйственные потребительские сбытовые (торговые) кооперативы" },
                    "непубличные акционерные общества": { "code": "12267", "name": "Непубличные акционерные общества" },
                    "союзы (ассоциации) кооперативов": { "code": "20605", "name": "Союзы (ассоциации) кооперативов" },
                    "муниципальные казенные учреждения": { "code": "75404", "name": "Муниципальные казенные учреждения" },
                    "муниципальные унитарные предприятия": { "code": "65243", "name": "Муниципальные унитарные предприятия" },
                    "адвокатские бюро": { "code": "20615", "name": "Адвокатские бюро" },
                    "сельскохозяйственные артели (колхозы)": { "code": "14153", "name": "Сельскохозяйственные артели (колхозы)" },
                }
                recData = copy.deepcopy(RECORDS_TEMPLATES.get(list_name))
                if list_name == "2. Реестр разрешений":
                    
                    if not TEST:
                        orgOGRN = row.get("ОГРН уполномоченного органа")
                        if pd.notna(orgOGRN) and orgOGRN.strip():
                            orgSearchParams = {"ogrn": str(orgOGRN.strip())}
                            unitSearch = get_unit(session, orgSearchParams, logger)
                            if unitSearch is not None:
                                unit = unitSearch.copy()
                                unit["id"] = unit.pop("_id")
                    recData["guid"] = generate_guid()
                    recData["parentEntries"] = "reestrpermitsReestr"
                    recData["unit"] = unit
                    recData["generalInformation"] = {
                        "Subject": row.get("Субъект РФ"),
                        "Disctrict": row.get("Муниципальный район/округ, городской округ или внутригородская территория")
                    }
                    recData["permission"] = {
                        "PermissionStatus": MAP_PERMISSION_STATUS[row.get("Статус разрешения").lower()],
                        "PermissionNumber": row.get("Номер разрешения"),
                        "PermissionStartDate": row.get("Дата выдачи разрешения"),
                        "administrationName": None,
                        "permissionEffectiveStartDate": row.get("Дата начала действия разрешения"),
                        "PermissionEndDate": row.get("Дата завершения действия разрешения"),
                        "permissionExtensionDate": row.get("Дата, до которой продлено действие разрешения"),
                        "reissuePermissionFile": None
                    }
                elif list_name == "3. Реестр рынков":
                    if not TEST:
                        orgOGRN = row.get("ОГРН уполномоченного органа")
                        if pd.notna(orgOGRN) and orgOGRN.strip():
                            orgSearchParams = {"ogrn": str(orgOGRN.strip())}
                            unitSearch = get_unit(session, orgSearchParams, logger)
                            if unitSearch is not None:
                                unit = unitSearch.copy()
                                unit["id"] = unit.pop("_id")
                    recData["guid"] = generate_guid()
                    recData["parentEntries"] = "reestrmarketReestr"
                    recData["unit"] = unit
                    recData["TotalInfo"] = {
                        "Subject":   row.get("Субъект РФ"),
                        "Disctrict": row.get("Муниципальный район/округ, городской округ или внутригородская территория")
                    }
                    recData["InfoRetailMarket"] = {
                        "RetailName": row.get("Наименование рынка"),
                        "PermissionStatus": MAP_PERMISSION_STATUS[row.get("Статус разрешения").lower()] if pd.notna(row.get("Статус разрешения")) else None,
                        "PermissionNumber": row.get("Номер разрешения") if pd.notna(row.get("Номер разрешения")) else None,
                        "PermissionStartDate": to_iso_date(row.get("Дата выдачи разрешения")) if pd.notna(row.get("Дата выдачи разрешения")) else None,
                        "PermissionEndDate": to_iso_date(row.get("Дата завершения действия разрешения")) if pd.notna(row.get("Дата завершения действия разрешения")) else None,
                        "marketType": MAP_MARKET_TYPE.get(row.get("Тип рынка").lower(), {}).get("code") if pd.notna(row.get("Тип рынка")) else None,
                        "marketSpecialization":    MAP_SPECIALIZATION.get(row.get("Специализация рынка").lower(), {}).get("code") if pd.notna(row.get("Специализация рынка")) else None,
                        "marketOtherSpecialization": row.get("Другая специализация рынка") if pd.notna(row.get("Другая специализация рынка")) else None,
                        "constituentFiles": None,
                        "GeoCoordinates": row.get("Геокоординаты точки, на которой расположен рынок") if pd.notna(row.get("Геокоординаты точки, на которой расположен рынок")) else None,
                        "marketArea": float(row.get("Площадь рынка, кв. м.")) if pd.notna(row.get("Площадь рынка, кв. м.")) else None,
                        "PlaceNumber": int(row.get("Число торговых мест, шт.")) if pd.notna(row.get("Число торговых мест, шт.")) else None,
                        "operationPeriod": MAP_OPERATION_PERIOD.get(row.get("Период действия разрешения").lower(), {"name": row.get("Период действия разрешения") }) if pd.notna(row.get("Период действия разрешения")) else None,
                        "marketOpeningTime": row.get("Основное время начала работы рынка") if pd.notna(row.get("Основное время начала работы рынка")) else None,
                        "marketClosingTime": row.get("Основное время окончания работы рынка") if pd.notna(row.get("Основное время окончания работы рынка")) else None,
                        "sanitaryDayOfMonth": row.get("Санитарный день месяца") if pd.notna(row.get("Санитарный день месяца")) else None,
                        "blockMarketAddress": [],
                        "blockCadNumber": [],
                        "cadsObjects": [],
                        "dayOnBlock": [],
                        "BlockDayOff": []
                    }
                    paUL = split_postal_address(row.get("Юридический адрес")) if pd.notna(row.get("Юридический адрес")) else {"postalCode": None, "fullAddress": None}
                    paFA = split_postal_address(row.get("Фактический адрес")) if pd.notna(row.get("Фактический адрес")) else {"postalCode": None, "fullAddress": None}
                    recData["InfoCompanyManagerMarket"] = {
                        "OperatorName":   row.get("Наименование оператора") if pd.notna(row.get("Наименование оператора")) else None,
                        "ShortNameUL":    row.get("Краткое наименование ЮЛ") if pd.notna(row.get("Краткое наименование ЮЛ")) else None,
                        "AddressUL":     { "postalCode": paUL["postalCode"], "fullAddress": paUL["fullAddress"] },
                        "AddressActual": { "postalCode": paFA["postalCode"], "fullAddress" : paFA["fullAddress"] },
                        "OperatorINN":    normStr(row.get("ИНН оператора")) if pd.notna(row.get("ИНН оператора")) else None,
                        "OperatorOGRN":   normStr(row.get("ОГРН оператора")) if pd.notna(row.get("ОГРН оператора")) else None,
                        "RykFIO":         normStr(row.get("ФИО руководителя")) if pd.notna(row.get("ФИО руководителя")) else None,
                        "OperatorNumber": normStr(row.get("Контактный номер оператора")) if pd.notna(row.get("Контактный номер оператора")) else None,
                        "OperatorEmail":  normStr(row.get("Адрес электронной почты оператора")) if pd.notna(row.get("Адрес электронной почты оператора")) else None,
                        "OrgStateForm": MAP_ORG_STATE_FORM.get(row.get("Организационно-правовая форма").lower(), {"name": row.get("Организационно-правовая форма") }) if pd.notna(row.get("Организационно-правовая форма")) else None
                    }

                if(TEST):
                    logger.info(f"TEST MODE: Структура для строки {i} | {json.dumps(recData, ensure_ascii=False)}")
                    continue

                recordURL = f"{BASE_URL}/api/v1/create/{recData['parentEntries']}"
                recordRes = api_request(session, logger, "post", recordURL, json=jsonable(recData))
                if not recordRes.ok:
                    logger.error(f"Ошибка при создании записи")
                    failLogger.info(i)
                    continue
                recordResJSON = recordRes.json()
                record_id = recordResJSON["_id"]
                record_guid = recordResJSON["guid"]
                if not TEST:
                    # Files pathes
                    files_pathes = []
                    fileExeption = False
                    if pd.notna(row.get("Разрешение на организацию розничного рынка")) and isinstance(row.get("Разрешение на организацию розничного рынка"), str) and row.get("Разрешение на организацию розничного рынка") != "":
                        fileIds = row.get("Разрешение на организацию розничного рынка").split(";")
                        for fileId in fileIds:
                            fileId_clean = fileId.replace("\n", "").replace("\r", "")
                            file_path = find_file_in_dir(FILES_DIR, fileId_clean)
                            if file_path:
                                logger.info(f"Найден файл: {os.path.basename(file_path)}")
                                files_pathes.append(file_path)
                            else:
                                logger.error(f"Файл не найден по шаблону: {fileId_clean}")
                                fileExeption = True
                                break
                    if fileExeption:
                        failLogger.info(i)
                        delete_from_collection(session, logger, recordResJSON)
                        logger.info(f"Ошибка при получении файла {i}")
                        continue
                    # Files upload
                    file_objects = []
                    exception_f_o = False
                    for file_p in files_pathes:
                        logger.info(f"Загрузка файла {file_p}")
                        file_object = upload_file(session, logger, file_p, "reestrpermitsReestr", record_id, entity_field_path="")
                        if file_object is None:
                            logger.error(f"Ошибка при загрузке файла {file_p}")
                            failLogger.info(i)
                            exception_f_o = True
                            break
                        file_objects.append(file_object)
                    if exception_f_o:
                        logger.info(f"Удаление данных по {i}")
                        for file_o in file_objects:
                            delete_file_from_storage(session, logger, file_o._id)
                            delete_from_collection(session, logger, recordResJSON)
                        failLogger.info(i)
                        logger.info(f"Завершено удаление данных по {i}")
                        continue
                    if len(file_objects) > 0:
                        updRecData = {
                            "_id": record_id,
                            "guid": record_guid,
                            "parentEntries": recData['parentEntries'],
                            "permission": {
                                "PermissionStatus": MAP_PERMISSION_STATUS[row.get("Статус разрешения").lower()],
                                "PermissionNumber": row.get("Номер разрешения"),
                                "PermissionStartDate": row.get("Дата выдачи разрешения"),
                                "administrationName": None,
                                "permissionEffectiveStartDate": row.get("Дата начала действия разрешения"),
                                "PermissionEndDate": row.get("Дата завершения действия разрешения"),
                                "permissionExtensionDate": row.get("Дата, до которой продлено действие разрешения"),
                                "reissuePermissionFile": file_objects[0]
                            }
                        }
                        recordUpdURL = f"{BASE_URL}/api/v1/update/{RECORDS_COLLECTION}?mainId={record_id}&guid={recordResJSON['guid']}"
                        recUpdRes = api_request(session, logger, "put", recordUpdURL, json=jsonable(updRecData))
                        if recUpdRes.status_code != requests.codes.ok:
                            logger.info(f"Удаление данных по {i}")
                            for file_o in file_objects:
                                delete_file_from_storage(session, logger, file_o["_id"])
                            delete_from_collection(session, logger, recordResJSON)
                            failLogger.info(i)
                            logger.info(f"Завершено удаление данных по {i}")
                            continue
                # Final log
                logger.info(f"Создана структура для строки {i} | _id записи - {record_id}")
                successLogger.info(json.dumps({"_id": record_id, "guid": record_guid, "parentEntries": RECORDS_COLLECTION}, ensure_ascii=False))
            logger.info(f"Завершена обработка листа: {list_name}")
        logger.info("Обработка файла завершена")

    except Exception as e:
        logger.error(f"Ошибка выполнения скрипта: {e}")


if __name__ == "__main__":
    main()