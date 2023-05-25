import re
import csv


def remove_fragments_with_highlighter(input_text):
    pattern = r"\\[^ ]*"
    result = re.sub(pattern, "", input_text)
    return result


def find_between(input_text, first, last):
    start = input_text.find(first) + len(first)
    end = input_text.find(last, start)
    return input_text[start:end] if start != -1 and end != -1 else None


def find_date(input_text):
    pattern = r"\b(\d{2}-\d{2}-\d{4})\b"  # date format dd-mm-rrrr
    match = re.search(pattern, input_text)
    if match:
        return match.group(1)
    else:
        return "date not found"


def create_dictionary(input_text):
    table = find_between(input_text, 'Data', 'Opis')
    table = table.replace('Cena m 2', '')
    date = find_date(table)
    repertorium = find_between(table, 'rep. ', ' ')
    parcel_number = find_between(table, repertorium + ' ', ' ')
    if input_text.find("Kondygnacja") == -1:
        tier = "brak danych"
    elif input_text.find("Liczba kondygnacji nadziemnych") > -1:
        tier = find_between(input_text, 'Kondygnacja:', 'Liczba')
    else:
        tier = find_between(input_text, 'Kondygnacja:', 'Ilosc')
    if input_text.find("Liczba kondygnacji nadziemnych") == -1:
        lkn = "brak danych"
    else:
        lkn = find_between(input_text, 'nadziemnych:  ', ' ')
    if input_text.find("Ilosc izb:") == -1:
        number_of_rooms = "brak danych"
    else:
        number_of_rooms = find_between(input_text, 'izb:', 'Lokalizacja')
    if input_text.find("Stan prawny:") == -1:
        legal_state = "brak danych"
    elif input_text.find("Rodzaj transakcji:") > -1:
        legal_state = find_between(input_text, 'prawny:', 'Rodzaj transakcji')
    else:
        legal_state = find_between(input_text, 'prawny:', 'Strony')
    if input_text.find("Rodzaj transakcji:") == -1:
        type_transaction = "brak danych"
    else:
        type_transaction = find_between(input_text, 'Rodzaj transakcji:', 'Strony')
    if input_text.find("pow. dzialki:") == -1:
        usage_area = "brak danych"
    else:
        usage_area = find_between(table, parcel_number, 'm')
    if usage_area == "brak danych":
        localization = find_between(input_text, 'Lokalizacja:', 'Data/')
    else:
        localization = find_between(input_text, 'Lokalizacja:', 'pow. dzialki')
    if input_text.find("Opis2") == -1:
        description1 = input_text[input_text.find('Opis:'):-1]
    else:
        description1 = find_between(input_text, 'Opis', 'Opis2')
    if input_text.find("Kondygnacja") == -1 and input_text.find("Liczba kondygnacji") > -1:
        type_location = find_between(input_text, 'lokalu:', 'Liczba kondygnacji')
    elif input_text.find("Kondygnacja") == -1 and input_text.find("Liczba kondygnacji") == -1:
        type_location = find_between(input_text, 'lokalu:', 'Ilosć izb')
    else:
        type_location = find_between(input_text, 'lokalu:', 'Kondygnacja')
    price = find_between(table, 'm 2', 'zl')
    dictionar = {
        "numer": find_between(input_text, 'nr', 'zr'),
        "zrodlo_notowania": find_between(input_text, 'notowania:', 'S' or 'R'),
        "stan_prawny": legal_state,
        "typ_transkacji": type_transaction,
        "strony": find_between(input_text, 'Strony:', 'Rodzaj'),
        "rodzaj_lokalu": type_location,
        "kondygnacje": tier,
        "lkn": lkn,  # liczba kondygnacji nadziemnych
        "liczba_izb": number_of_rooms,
        "lokalizacja": localization,
        "data": date,
        "repertorium": repertorium,
        "numer_dzialki": parcel_number,
        "pow_uzytkowa": usage_area,
        "cena": price,
        "cena_za_metr2": find_between(table, 'zl', 'zl/m 2'),
        "opis1": description1,
        "opis2": input_text[input_text.find('Opis2KW'):-1],
    }
    return dictionar


y = input("Podaj nazwe pliku: ")
x = int(input("Podaj liczbe transakcji: "))
with open(y, 'r') as file:
    text = file.read()
text = text.replace("\\u322?", "l")
text = text.replace("\\u347?", "s")
text = text.replace("\\u377?", "z")
text = text.replace("\\u243?", "o")
text = text.replace("\\u263?", "ć")
text = text.replace("\\u380?", "z")
text = text.replace("\\u321?", "l")
text = text.replace("\\u261?", "a")
text = text.replace("\\u281?", "e")
text = text.replace("ć", "c")
text = text.replace("ł", "l")
text = text.replace("\\u160?", "")
text = remove_fragments_with_highlighter(text)
transaction = []
transaction_dict = []
for i in range(x-1):
    trans = find_between(text, 'Transakcja', 'Transakcja')
    trans = "Transakcja"+trans
    text = text.replace(trans, '')
    transaction.append(trans)
    dicti = create_dictionary(transaction[i])
    transaction_dict.append(dicti)
trans = find_between(text, 'Transakcja', '}')
transaction.append("Transakcja"+trans)
dicti = create_dictionary(trans)
transaction_dict.append(dicti)
keys = transaction_dict[0].keys()
with open('export_rtf.csv', 'w') as output_file:
    dict_writer = csv.DictWriter(output_file, keys, delimiter=';')
    dict_writer.writeheader()
    dict_writer.writerows(transaction_dict)
