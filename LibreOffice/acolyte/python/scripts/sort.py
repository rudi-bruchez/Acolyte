import uno
import unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.beans import PropertyValue
import unicodedata
import re
import xml.etree.ElementTree as ET
from io import StringIO

ID_EXTENSION = 'org.indoiranian.wordlist.sorter'
SERVICE = ('com.sun.star.task.Job',)

SUBSCRIPT_HOOKS = ["'", "’", "ʾ"]  # Characters to be treated as subscript_hook

def normalize_khotanese(word):
    word = unicodedata.normalize('NFD', word)
    char_map = {
        '\u0328': '', '\u0105': 'a', '\u012F': 'i', '\u0173': 'u', '\u0119': 'e', '\u01EB': 'o',
        '\u0104': 'A', '\u012E': 'I', '\u0172': 'U', '\u0118': 'E', '\u01EA': 'O',
    }
    normalized = ''.join(char_map.get(c, c) for c in word)
    return unicodedata.normalize('NFC', normalized)

def normalize_special_vowels(word):
    special_vowels = {
        'a\u0304\u0306': 'a_l+b', '\u0101\u0306': 'a_l+b',
        'i\u0304\u0306': 'i_l+b', '\u012b\u0306': 'i_l+b',
        'u\u0304\u0306': 'u_l+b', '\u016b\u0306': 'u_l+b',
        'A\u0304\u0306': 'A_l+b', '\u0100\u0306': 'A_l+b',
        'I\u0304\u0306': 'I_l+b', '\u012A\u0306': 'I_l+b',
        'U\u0304\u0306': 'U_l+b', '\u016A\u0306': 'U_l+b',
    }
    for combo, replacement in special_vowels.items():
        word = word.replace(combo, replacement)
    return word

def sort_words_in_document(language):
    try:
        localContext = uno.getComponentContext()
        smgr = localContext.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", localContext)
        model = desktop.getCurrentComponent()

        if hasattr(model, "Sheets"):
            sort_calc_document(model, language)
        elif hasattr(model, "Text"):
            sort_writer_document_xml(model, language)
        else:
            raise Exception("This document is neither a Calc spreadsheet nor a Writer document.")
    except Exception as e:
        print(f"Error during sorting: {e}")

def sort_calc_document(model, language):
    try:
        sheet = model.CurrentController.ActiveSheet
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(True)
        data = cursor.getDataArray()

        if not data or len(data) == 0:
            print("No data found in the sheet.")
            return

        sorted_data = sorted(data, key=lambda row: custom_sort(str(row[0]), language))
        target_range = sheet.getCellRangeByPosition(0, 0, len(data[0])-1, len(data)-1)
        target_range.setDataArray(sorted_data)
        print("Calc document sorted successfully.")
    except Exception as e:
        print(f"Error in sorting Calc document: {e}")

def sort_writer_document_xml(model, language):
    try:
        print("Starting sort_writer_document_xml")

        text = model.Text

        # Get all paragraphs
        paragraphs = []
        enum = text.createEnumeration()
        while enum.hasMoreElements():
            paragraph = enum.nextElement()
            if paragraph.supportsService("com.sun.star.text.Paragraph"):
                paragraphs.append(paragraph)

        # Extract text and formatting from paragraphs
        paragraph_data = []
        for p in paragraphs:
            paragraph_content = []
            portion_enum = p.createEnumeration()
            while portion_enum.hasMoreElements():
                text_portion = portion_enum.nextElement()
                if text_portion.supportsService("com.sun.star.text.TextPortion"):
                    content = text_portion.getString()
                    if content.strip():
                        properties = {
                            'CharPosture': text_portion.CharPosture,
                            'CharWeight': text_portion.CharWeight,
                            'CharColor': text_portion.CharColor,
                            'CharUnderline': text_portion.CharUnderline,
                            'CharBackColor': text_portion.CharBackColor,
                            'CharFontName': text_portion.CharFontName,  # Aggiungiamo il nome del font
                            'CharHeight': text_portion.CharHeight,  # Aggiungiamo la dimensione del carattere
                            'CharLocale': text_portion.CharLocale,  # Aggiungiamo la lingua del carattere
                        }
                        paragraph_content.append((content, properties))
            if paragraph_content:
                paragraph_data.append(paragraph_content)

        print(f"Number of paragraphs: {len(paragraph_data)}")
        print("Paragraph content:")
        for p in paragraph_data:
            print([f"{content} ({props})" for content, props in p])

        # Sort paragraphs
        sorted_paragraphs = sorted(paragraph_data, key=lambda x: custom_sort(''.join(portion[0] for portion in x), language))

        # Clear existing content
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        cursor.setString("")

        # Insert sorted paragraphs
        for paragraph in sorted_paragraphs:
            for content, properties in paragraph:
                # Insert the content
                text.insertString(cursor, content, False)

                # Select the inserted text
                cursor.goLeft(len(content), True)

                # Apply stored properties
                for prop, value in properties.items():
                    cursor.setPropertyValue(prop, value)

                # Move cursor to end of inserted text
                cursor.goRight(len(content), False)

            # Move to next line
            text.insertControlCharacter(cursor, 0, False)  # 0 is the constant for newline

        print("Writer document sorted successfully.")
    except Exception as e:
        print(f"Error in sorting Writer document: {e}")
        import traceback
        traceback.print_exc()

def custom_sort(word, language):
    if not word:
        return ("", "", "", "")

    superscript_match = re.match(r'^([⁰¹²³⁴⁵⁶⁷⁸⁹]+°?)', word)
    superscript = superscript_match.group(1) if superscript_match else ''

    word_without_superscript = re.sub(r'^[⁰¹²³⁴⁵⁶⁷⁸⁹]+°?', '', word)

    number = extract_number(word_without_superscript)
    word_without_number = remove_leading_numbers(word_without_superscript)

    main_word = word_without_number.rstrip('‑-')
    has_hyphen = word_without_number.endswith(('‑', '-'))

    lowercase_word = to_lowercase_special(main_word)
    preprocessed_word = preprocess_word(lowercase_word, language)

    mapping = get_mapping(language)

    transformed_word = ""
    i = 0
    while i < len(preprocessed_word):
        if i < len(preprocessed_word) - 4 and preprocessed_word[i:i+5] in mapping:
            transformed_word += mapping[preprocessed_word[i:i+5]]
            i += 5
        elif i < len(preprocessed_word) - 3 and preprocessed_word[i:i+4] in mapping:
            transformed_word += mapping[preprocessed_word[i:i+4]]
            i += 4
        elif i < len(preprocessed_word) - 2 and preprocessed_word[i:i+3] in mapping:
            transformed_word += mapping[preprocessed_word[i:i+3]]
            i += 3
        elif i < len(preprocessed_word) - 1 and preprocessed_word[i:i+2] in mapping:
            transformed_word += mapping[preprocessed_word[i:i+2]]
            i += 2
        elif preprocessed_word[i] in mapping:
            transformed_word += mapping[preprocessed_word[i]]
            i += 1
        else:
            transformed_word += preprocessed_word[i]
            i += 1

    superscript_value = ''.join([str('⁰¹²³⁴⁵⁶⁷⁸⁹'.index(c)) for c in superscript if c in '⁰¹²³⁴⁵⁶⁷⁸⁹'])
    superscript_value = superscript_value.zfill(5)

    hyphen_value = '1' if has_hyphen else '0'

    return (transformed_word, superscript_value, hyphen_value, number)

def extract_number(word):
    match = re.match(r'^([0-9¹²³]+)', word)
    return match.group(1) if match else ''

def remove_leading_numbers(word):
    return re.sub(r'^[0-9¹²³°]+\s*', '', word)

def remove_parentheses(word):
    return re.sub(r'\([^)]*\)', '', word)

def to_lowercase_special(word):
    lowercase_map = {
        'A': 'a', 'Ā': 'ā', 'I': 'i', 'Ī': 'ī', 'U': 'u', 'Ū': 'ū',
        'Ṛ': 'ṛ', 'Ṝ': 'ṝ', 'Ḷ': 'ḷ', 'E': 'e', 'O': 'o',
        'K': 'k', 'G': 'g', 'Ṅ': 'ṅ', 'C': 'c', 'J': 'j', 'Ñ': 'ñ',
        'Ṭ': 'ṭ', 'Ḍ': 'ḍ', 'Ṇ': 'ṇ', 'T': 't', 'D': 'd', 'N': 'n',
        'P': 'p', 'B': 'b', 'M': 'm', 'Y': 'y', 'R': 'r', 'L': 'l',
        'V': 'v', 'Ś': 'ś', 'Ṣ': 'ṣ', 'S': 's', 'H': 'h',
        'Ai': 'ai', 'Au': 'au', 'Ei': 'ei',
        'Ph': 'ph', 'Th': 'th', 'Ch': 'ch', 'Kh': 'kh', 'Gh': 'gh',
        'Gy': 'gy', 'Ky': 'ky',
        'A_l+b': 'a_l+b', 'I_l+b': 'i_l+b', 'U_l+b': 'u_l+b',
        'A\u0304\u0306': 'a_l+b', '\u0100\u0306': 'a_l+b',
        'I\u0304\u0306': 'i_l+b', '\u012A\u0306': 'i_l+b',
        'U\u0304\u0306': 'u_l+b', '\u016A\u0306': 'u_l+b',
    }

    for multi, lower in lowercase_map.items():
        if len(multi) > 1 or '_' in multi or r'\u' in multi:
            word = word.replace(multi, lower)

    return ''.join(lowercase_map.get(c, c.lower()) for c in word)

def preprocess_word(word, language):
    word = re.sub(r"[\(\)\[\]\{\}\*°:‐]", "", word)
    word = remove_parentheses(word)
    word = normalize_special_vowels(word)

    if language == "sanskrit":
        return preprocess_sanskrit_word(word)
    elif language == "khotanese":
        word = normalize_khotanese(word)
        word = word.replace('r̥', 'rä')
        word = word.replace('ṃ', '')
    return word.strip()

def preprocess_sanskrit_word(word):
    preprocessed_word = ""
    i = 0
    while i < len(word):
        if word[i] == 'ṃ':
            if i < len(word) - 1 and word[i + 1] in 'kgcṭtdpb':
                nasal_replacements = {'k': 'ṅ', 'g': 'ṅ', 'c': 'ñ', 'ṭ': 'ṇ', 't': 'n', 'd': 'n', 'p': 'm', 'b': 'm'}
                preprocessed_word += nasal_replacements.get(word[i + 1], 'ṃ')
            else:
                preprocessed_word += 'ṃ'
        else:
            preprocessed_word += word[i]