import re
import xlsxwriter

from datetime import datetime, timedelta


class JsonStructureAnalizer:
    def __init__(self):
        self.names = {}
        self.max_deep = 0

    def get_rows_and_names(self, rec):
        if isinstance(rec, dict):
            self.__get_rows_and_names_from_dict(rec, "", 0)
        elif isinstance(rec, list):
            self.__get_rows_and_names_from_list(rec, "", 0)
        else:
            self.names[""] = 1

    def __get_rows_and_names_from_dict(self, rec: dict, name, deep):
        deep += 1
        if self.max_deep < deep:
            self.max_deep = deep
        # Если есть хотя бы одна запись не являющаяся списком или словарём в этом словаре, значит будет существовать
        # колонка с соответствующим именем.
        for key, value in rec.items():
            new_name = f"{name}_{key}".strip("_") if name else key

            if isinstance(value, dict):
                self.__get_rows_and_names_from_dict(value, new_name, deep)
            elif isinstance(value, list):
                self.__get_rows_and_names_from_list(value, new_name, deep)
            else:
                if new_name not in self.names:
                    self.names[new_name] = 0
                self.names[new_name] += 1

    def __get_rows_and_names_from_list(self, rec: list, name, deep):
        deep += 1
        if self.max_deep < deep:
            self.max_deep = deep
        # Если есть хотя бы одна запись не являющаяся списком или словарём в этом списке, значит будет существовать
        # колонка соответствующим именем.
        for value in rec:
            if isinstance(value, dict):
                self.__get_rows_and_names_from_dict(value, name, deep)
            elif isinstance(value, list):
                self.__get_rows_and_names_from_list(value, name, deep)
            else:
                if name not in self.names:
                    self.names[name] = 0
                self.names[name] += 1


class WorksheetWriter:
    ifd = r"(?P<date>(?P<year>\d{2,4})[-_.\/\\](?P<month>\d{1,2})([-_.\/\\](?P<day>\d{1,2}))?)?"
    ift = (
        r"(?P<time>(?P<hour>\d{1,2}):(?P<minute>\d{1,2})(:(?P<second>\d{1,2}(\.\d+)?))?(?P<timezone>(?P<tz_sign>\+|-)?"
        r"(?P<tz_hour>\d{1,2})((\.|:)?(?P<tz_minute>\d{1,2})))?)?"
    )
    input_formats_date_time = r"^" + ifd + r" *[ tTzZ] *" + ift + r"$"
    input_formats_time_date = r"^" + ift + r" *" + ifd + r"$"
    input_formats_date = r"^" + ifd + r"$"
    input_formats_time = r"^" + ift + r"$"

    def __init__(self, output_xlsx_file_name):
        self.__workbook = None
        self.__worksheet = None
        self.__current_recs = []
        self.__names = {}
        self.__deep = 0
        self.__freeze = {}
        self.__dinamic = {}
        self.__row = 0
        self.__col = 0
        self.__name_size = []
        self.output_xlsx_file_name = output_xlsx_file_name

    def __enter__(self):
        self.__workbook = xlsxwriter.Workbook(self.output_xlsx_file_name)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.__workbook.close()

    def set_cols_size(self):
        for col in range(len(self.__names)):
            self.__worksheet.set_column(col, col, min(200, self.__name_size[col]))

    def write(self, data):
        if isinstance(data, list):
            out_data = {}
            list_default = []
            num = 1
            for value in data:
                if isinstance(value, list):
                    out_data[f"Лист {num}"] = value
                    num += 1
                else:
                    list_default.append(value)
            if list_default:
                out_data["Лист 0"] = list_default
            data = out_data

        structure = {}
        for key, value in data.items():
            analizer = JsonStructureAnalizer()
            analizer.get_rows_and_names(value)
            structure[key] = {
                "": analizer,
                "rows": sum(analizer.names.values()),
                "cols": len(analizer.names),
                "deep": analizer.max_deep,
                "names": {
                    tuple(name.split("_")): number
                    for number, name in enumerate(sorted(analizer.names))
                },
            }

        for list_name, rec in data.items():
            self.__worksheet = self.__workbook.add_worksheet(list_name)

            self.__names = structure[list_name]["names"]
            self.__deep = max(map(len, self.__names))
            self.__col = structure[list_name]["cols"]
            self.__name_size = [10] * self.__col

            self.__write_headers()
            self.__write(rec)
            self.set_cols_size()

    def __write_headers(self):
        self.__row = 0
        self.__col = 0
        for deep in range(self.__deep):
            self.__current_recs = [
                [
                    names[deep] if len(names) > deep else None
                    for names in self.__names.keys()
                ]
            ]
            bold = self.__workbook.add_format({"bold": True})
            self.__write_current_recs(style=bold)

    def __write_one_cell(self, row, col, value, style):
        def match_datatime(match):
            try:
                year = int(match.group("year"))
                month = int(match.group("month"))
                try:
                    day = int(match.group("day"))
                except IndexError:
                    day = 1
                except TypeError:
                    day = 1
            except IndexError:
                year = 1970
                month = 1
                day = 1
            except TypeError:
                year = 1970
                month = 1
                day = 1

            try:
                hour = int(match.group("hour"))
                minute = int(match.group("minute"))
                try:
                    f_second = float(match.group("second"))
                except IndexError:
                    f_second = 0
                except TypeError:
                    f_second = 0
            except TypeError:
                hour = 0
                minute = 0
                f_second = 0
            except IndexError:
                hour = 0
                minute = 0
                f_second = 0

            second = int(f_second)
            microsecond = int(1e6 * (f_second - second))

            try:
                tz_hour = int(match.group("tz_hour"))
                try:
                    tz_minute = int(match.group("tz_minute"))
                except TypeError:
                    tz_minute = 0
                except IndexError:
                    tz_minute = 0
                try:
                    tz_sign = match.group("tz_sign")
                    tz_sign = -1 if tz_sign == "-" else 1
                except IndexError:
                    tz_sign = 1
                except TypeError:
                    tz_sign = 1
            except IndexError:
                tz_hour = 0
                tz_minute = 0
                tz_sign = 1
            except TypeError:
                tz_hour = 0
                tz_minute = 0
                tz_sign = 1

            dt = datetime(year, month, day, hour, minute, second, microsecond)
            dt += timedelta(hours=tz_sign * tz_hour, minutes=tz_sign * tz_minute)
            return dt

        if isinstance(value, str):
            # URL block
            pattern = r"^(?P<url>(http|ftp)(s)?:\/(\/[А-Яа-яЁёA-Za-z0-9-_\.]+)+(\/)?)$"
            match = re.search(pattern, value)
            if match and match.group("url"):
                self.__worksheet.write_url(
                    row,
                    col,
                    url=match.group("url"),
                    cell_format=style,
                    string=match.group("url"),
                )
                return
            pattern = r"^(?P<text>.+) (?P<url>(http|ftp)(s)?:\/(\/[А-Яа-яЁёA-Za-z0-9-_\.]+)+(\/)?)$"
            match = re.search(pattern, value)
            if match and match.group("url"):
                self.__worksheet.write_url(
                    row,
                    col,
                    url=match.group("url"),
                    cell_format=style,
                    string=match.group("text"),
                )
                return
            pattern = r"^(?P<url>(http|ftp)(s)?:\/(\/[А-Яа-яЁёA-Za-z0-9-_\.]+)+(\/)?) (?P<text>.+)$"
            match = re.search(pattern, value)
            if match and match.group("url"):
                self.__worksheet.write_url(
                    row,
                    col,
                    url=match.group("url"),
                    cell_format=style,
                    string=match.group("text"),
                )
                return
            pattern = r"(?P<text>.+ (?P<url>(http|ftp)(s)?:\/(\/[А-Яа-яЁёA-Za-z0-9-_\.]+)+(\/))? .+)"
            match = re.search(pattern, value)
            if match and match.group("url"):
                self.__worksheet.write_url(
                    row,
                    col,
                    url=match.group("url"),
                    cell_format=style,
                    string=match.group("text"),
                )
                return
            # DateTime block
            for pattern in list(
                map(
                    re.compile,
                    [
                        self.input_formats_date_time,
                        self.input_formats_time_date,
                        self.input_formats_date,
                        self.input_formats_time,
                    ],
                )
            ):
                match = re.search(pattern, value)
                if match and len(match.group(0)) > 4:
                    try:
                        dt = match_datatime(match)
                        dtf = {"num_format": "yyyy-mm-dd hh:mm:ss.000"}
                        dtf.update(style)
                        format = self.__workbook.add_format(dtf)
                        self.__worksheet.write_datetime(row, col, dt, format)
                        return
                    except:
                        pass
        self.__worksheet.write(row, col, value, style)

    def __write_current_recs(self, style=None):
        if style is None:
            style = {}
        for current_rec in self.__current_recs:
            for col, value in enumerate(current_rec):
                if value is None:
                    continue
                self.__name_size[col] = max(self.__name_size[col], len(str(value)))
                self.__write_one_cell(self.__row, col, value, style)
            self.__row += 1

    def __write(self, rec):
        if not isinstance(rec, list):
            rec = [rec]
        for rec in rec:
            self.__current_recs = []
            if isinstance(rec, dict):
                static, dinamic = self.__form_rows_from_dict(rec, tuple(), {}, {})
                # Всё содержимое из freeze копируем как есть.
                if dinamic:
                    self.__form_rows_from_list(static, dinamic)
                else:
                    self.__current_recs.append(
                        [static.get(name) for name in self.__names]
                    )
            else:
                if not isinstance(rec, list):
                    rec = [rec]
                self.__current_recs.append(rec)
            self.__write_current_recs()

    def __form_rows_from_dict(
        self, rec: dict, last_key: tuple, static: dict, dinamic: dict
    ):
        check_dict = {
            key_0 if isinstance(key_0, tuple) else tuple((key_0,)): value_0
            for key_0, value_0 in rec.items()
        }
        new_static = static.copy()
        # Перебираю пока в словаре на текущем уровне не останется других словарей,
        # должны остаться лишь списки и значения.
        while check_dict:
            _key, value = check_dict.popitem()
            key = tuple((*last_key, *_key)) if last_key else _key
            if isinstance(value, dict):
                for new_key, value in value.items():
                    if new_key == "":
                        check_dict[key] = value
                    else:
                        check_dict[tuple((*key, new_key))] = value
            elif isinstance(value, list):
                dinamic[key] = value[::-1]
            else:
                new_static[key] = value
        return new_static, dinamic

    def __form_rows_from_list(self, static: dict, dinamic: dict):
        # Перебираю каждый член из dinamic.
        new_data_found = True
        while new_data_found:
            new_dinamic = {}
            new_static = {}

            add_static = static.copy()

            new_data_found = False
            for key, value in dinamic.items():
                if not value:
                    add_static[key] = None
                    continue

                pop_value = value.pop()

                if isinstance(pop_value, list):
                    # Появился новый список, добавляю его в новый лист откуда буду брать значения.
                    new_dinamic[key] = pop_value[::-1]
                    new_data_found = True

                elif isinstance(pop_value, dict):
                    # Добавляю значение по умолчанию, если оно есть.
                    if "" in pop_value:
                        add_static[key] = pop_value.pop("")
                        new_data_found = True
                    else:
                        add_static[key] = None
                    # Если в словаре остались иные значения, добавляю словарь в список для будущего добавления в
                    # статику.
                    if pop_value:
                        new_static[key] = pop_value
                        new_data_found = True
                else:
                    add_static[key] = pop_value
                    new_data_found = True

            if new_data_found:
                if new_static:
                    for key, value in new_static.items():
                        _static, _dinamic = self.__form_rows_from_dict(
                            value, key, add_static, new_dinamic
                        )
                        add_static.update(_static)
                        new_dinamic.update(_dinamic)
                if new_dinamic:
                    self.__form_rows_from_list(add_static, new_dinamic)
                else:
                    self.__current_recs.append(
                        [add_static.get(name) for name in self.__names]
                    )
