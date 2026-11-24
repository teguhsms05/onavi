"""Fungsi Excel yang belum ada di pycel, dipasang sebagai plugin ExcelCompiler (nama fungsi = nama Excel, huruf kecil)."""
# Import ber-alias: pycel memperlakukan setiap nama di modul ini sebagai fungsi Excel (mis. `date` -> DATE)
import datetime as _dt

EPOCH = _dt.date(1899, 12, 30)  # tanggal serial Excel 0


def datedif(start_date, end_date, unit):
    """DATEDIF(mulai, akhir, satuan) seperti Excel: satuan Y, M, D, MD, YM, YD. Mulai > akhir -> #NUM!."""
    try:
        start = EPOCH + _dt.timedelta(days=int(start_date or 0))
        end = EPOCH + _dt.timedelta(days=int(end_date or 0))
    except (TypeError, ValueError):
        return "#VALUE!"
    if start > end:
        return "#NUM!"
    unit = str(unit).upper()
    months = (end.year - start.year) * 12 + end.month - start.month - (end.day < start.day)
    if unit == "Y":
        return months // 12
    if unit == "M":
        return months
    if unit == "D":
        return (end - start).days
    if unit == "YM":
        return months % 12
    if unit == "MD":
        if end.day >= start.day:
            return end.day - start.day
        prev_month_end = end.replace(day=1) - _dt.timedelta(days=1)
        return (prev_month_end - prev_month_end.replace(day=min(start.day, prev_month_end.day))).days + end.day
    if unit == "YD":
        try:
            anniv = start.replace(year=end.year)
        except ValueError:  # 29 Feb di tahun bukan kabisat
            anniv = _dt.date(end.year, 3, 1)
        if anniv > end:
            anniv = anniv.replace(year=end.year - 1) if not (start.month == 2 and start.day == 29) else _dt.date(end.year - 1, 3, 1)
        return (end - anniv).days
    return "#NUM!"
