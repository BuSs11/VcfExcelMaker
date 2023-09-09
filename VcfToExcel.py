import vobject
from openpyxl import Workbook
import quopri

def decode_quoted_printable(text):
    try:
        return quopri.decodestring(text).decode('utf-8')
    except:
        return text

def vcf_to_excel(vcf_file, output_excel):
    with open(vcf_file, 'r', encoding='utf-8') as f:
        vcard_data = f.readlines()

    vcards = []
    current_vcard = None
    for line in vcard_data:
        line = line.strip()

        # QUOTED-PRINTABLE decoding
        if "=QUOTED-PRINTABLE:" in line:
            parts = line.split("=QUOTED-PRINTABLE:")
            if len(parts) == 2:
                line = parts[0] + ":" + decode_quoted_printable(parts[1])

        if line.startswith("BEGIN:VCARD"):
            current_vcard = [line]
        elif line.startswith("END:VCARD") and current_vcard is not None:
            current_vcard.append(line)
            vcards.append(current_vcard)
            current_vcard = None
        elif current_vcard is not None:
            current_vcard.append(line)

    # 이름 순으로 정렬
    def vcard_sort_key(vcard_lines):
        try:
            vcard = vobject.readOne("\n".join(vcard_lines))
            if hasattr(vcard, 'fn'):
                return vcard.fn.value
            return ""
        except:
            return ""
    
    vcards.sort(key=vcard_sort_key)

    # 엑셀 워크북 및 워크시트 초기화
    wb = Workbook()
    ws = wb.active
    ws.append(["이름", "라인"])
    
    for vcard_lines in vcards:
        try:
            vcard = vobject.readOne("\n".join(vcard_lines))
            for vcard_line in vcard_lines:
                ws.append([vcard.fn.value, vcard_line])
        except Exception as e:
            print("Error parsing vcard:", e)
            for vcard_line in vcard_lines:
                ws.append([None, vcard_line])

    # 엑셀 파일로 저장
    wb.save(output_excel)

vcf_file = ""
output_excel = ""
vcf_to_excel(vcf_file, output_excel)
