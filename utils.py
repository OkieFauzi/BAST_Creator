from datetime import datetime

def spell_number(number):
    units = ["nol", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan"]
    teens = ["sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas"]
    tens = ["", "sepuluh", "dua puluh", "tiga puluh", "empat puluh", "lima puluh", "enam puluh", "tujuh puluh", "delapan puluh", "sembilan puluh"]

    if number < 10:
        return units[number]
    elif 10 <= number < 20:
        return teens[number - 10]
    elif 20 <= number < 100:
        return tens[number // 10] + (" " + spell_number(number % 10) if number % 10 != 0 else "")
    elif 100 <= number < 1000:
        return units[number // 100] + " ratus" + (" " + spell_number(number % 100) if number % 100 != 0 else "")
    elif 1000 <= number < 1000000:
        return spell_number(number // 1000) + " ribu" + (" " + spell_number(number % 1000) if number % 1000 != 0 else "")
    else:
        return "Angka terlalu besar untuk dieja"

def spell_date(date_input, return_format="full"):
    days_list = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    months_list = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]

    # If input is a string, convert it to datetime
    if isinstance(date_input, str):
        date_obj = datetime.strptime(date_input, "%Y-%m-%d")
    elif isinstance(date_input, datetime):
        date_obj = date_input
    else:
        return "Format tanggal tidak valid"

    day_name = days_list[date_obj.weekday()]
    day_number = date_obj.day
    month_name = months_list[date_obj.month - 1]
    year_number = date_obj.year

    day_spelled = spell_number(day_number)
    year_spelled = spell_number(year_number)

    if return_format == "day":
        return day_name
    elif return_format == "date":
        return day_spelled
    elif return_format == "month":
        return month_name
    elif return_format == "year":
        return year_spelled
    elif return_format == "full":
        return f"{day_name}, {day_spelled} {month_name} {year_spelled}"
    else:
        return "Format keluaran tidak valid"