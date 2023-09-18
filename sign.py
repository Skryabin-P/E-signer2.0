
import argparse
import PyPDF2
import re
import sys
import datetime
from cryptography.hazmat import backends
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.x509.oid import NameOID
from endesive import pdf
from datetime import datetime
signature_string = lambda organization, date, country : (organization + '\nDATE: '+ date)



def load_pfx(file_path, password):
    """ Function to load pkcs12 object from the given password protected pfx file."""

    with open(file_path, 'rb') as fp:
        return pkcs12.load_key_and_certificates(fp.read(), password.encode(), backends.default_backend())



OID_NAMES = {
    NameOID.COMMON_NAME: 'CN',
    NameOID.COUNTRY_NAME: 'C',
    NameOID.DOMAIN_COMPONENT: 'DC',
    NameOID.EMAIL_ADDRESS: 'E',
    NameOID.GIVEN_NAME: 'G',
    NameOID.LOCALITY_NAME: 'L',
    NameOID.ORGANIZATION_NAME: 'O',
    NameOID.ORGANIZATIONAL_UNIT_NAME: 'OU',
    NameOID.SURNAME: 'SN'
}

def get_rdns_names(rdns):
    # parsing data from .pfx certificate
    names = {}
    for oid in OID_NAMES:
        names[OID_NAMES[oid]] = ''
    for rdn in rdns:
        for attr in rdn._attributes:
            if attr.oid in OID_NAMES:
                names[OID_NAMES[attr.oid]] = attr.value
    return names

def beauty_fingerprint(string:str):
    string = string.swapcase()
    # string = " ".join(string[i:i+2] for i in range(0, len(string), 2))
    new_str = ''
    for i in range(0,len(string),2):
        new_str += string[i:i + 2] + ' '
        if i == 28:
            new_str+='\n'
    return new_str

def sign(func_position: str, pfx_certificate, password, input_file, dest):
    try:
        first_time = True

        # Load the PKCS12 object from the pfx file
        p12pk, p12pc, p12oc = load_pfx(pfx_certificate, password)
        a = p12pc.fingerprint(p12pc.signature_hash_algorithm).hex()
        fingerprint = beauty_fingerprint(a)
        # since we sign files with each certificate separately , after first signing we save signed file in the destination save folder
        # and after that we should read that signed file to sign with another certificates
        if not first_time:
            input_file = dest
        names = get_rdns_names(p12pc.subject.rdns)
        pdf_f = open(input_file, 'rb')
        pdfread = PyPDF2.PdfFileReader(pdf_f)
        num_pages = pdfread.getNumPages()
        pages = num_pages
        start = num_pages - 1

        dest = dest if dest else input_file
        if all:
            start = 0  # if checked "sign all pages" then start signing from first page otherwise from the last
        for page in range(start, pages):
            if not first_time:
                input_file = dest
            date = datetime.now()
            date = date.strftime(f'%Y%m%d%H%M%S+03\'00\'')

            pdf_f = open(input_file, 'rb')
            pdfread = PyPDF2.PdfFileReader(pdf_f)

            fields = pdfread.get_fields()
            # collect name fields in document
            if str(type(fields)) != "<class 'NoneType'>":
                signs = fields.keys()
            else:
                signs = ['nothing']
            i = 1
            # check how many our signs on page
            for sign in signs:
                for t in range(1, 4):
                    mask = f"Signature{t}_p{page + 1}s"
                    if mask in sign:
                        i += 1

            if i == 1:
                plus_coord = 0
            else:
                plus_coord = 250 * (i - 1)

            # parameters of the sign like size, align, place coords, etc
            # signature = signature_string(names['CN'], date, names['C'])

            first = names["CN"][0].lower()
            signature_date = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            if first == 'ш':
                name = names["CN"][1:]
                dct = {
                    "aligned": 0,
                    "sigflags": 3,
                    "sigflagsft": 132,
                    "sigpage": pages - 1,
                    'sigandcertify': True,
                    "auto_sigfield": False,
                    # "sigandcertify": False,
                    "signaturebox": (5 + plus_coord, 5, 255 + plus_coord, 80),
                    "signform": False,
                    "sigfield": f"Signature{i}_p{page + 1}s",
                    # "signature_appearance": {
                    #     'background': [1, 1, 1],
                    #     # 'icon': '../signature_test.png',
                    #     'outline': [0.2, 0.3, 0.5],
                    #     'border': 2,
                    #     'labels': False,
                    #     'display': 'CN,contact,date'.split(','),
                    #     # 'display': signature,
                    #     },

                    "signature_manual": [
                        ['background', 1, 0, 1],
                        # ['fill_colour', 1, 0, 0],
                        # ['rect_fill', 0, 50, 108, 150], ,['outline',1, 1, 0.5],
                        ['stroke_color', 0, 0, 1],  # border color
                        ['border', 1],
                        ['fill_colour', 0, 0, 1],  # text color
                        ['text_box', f'Документ подписан электронно', 'default', 0, 60, 270, 10, 7, True, 'center',
                         'top'],
                        ['text_box', f'Подписал(ФИО):\nEmail:\nДолжность:\nВремя подписания\nСертификат:',

                         # font  *[bounding box], size, wrap, align, baseline, spacing
                         'default', 5, 15, 80, 40, 7, True, 'left', 'top'],
                        ['text_box', f'{name}',
                         'ttf_font', 97.5, 15, 270, 40, 7, True, 'left', 'top'],
                        ['text_box', f'\n {names["E"]}\n{func_position}\n{signature_date}\n{fingerprint}',
                         'ttf_font', 90, 15, 270, 40, 7, True, 'left', 'top'],
                        ['text_box', f'{first}',
                         'ttf_font', 90, 18.5, 270, 40, 10, True, 'left', 'top']

                    ],

                    "contact": names['E'],
                    "location": "Russia",
                    "signingdate": date,
                    "reason": "Подписание",
                    # "signature": signature,
                    "password": password,

                }
            else:
                name = names["CN"]

                dct = {
                    "aligned": 0,
                    "sigflags": 3,
                    "sigflagsft": 132,
                    "sigpage": pages - 1,
                    'sigandcertify': True,
                    "auto_sigfield": False,
                    # "sigandcertify": False,
                    "signaturebox": (40 + plus_coord, 40, 290 + plus_coord, 120),
                    "signform": False,
                    "sigfield": f"Signature{i}_p{page + 1}s",
                    # "signature_appearance": {
                    #     'background': [1, 1, 1],
                    #     # 'icon': '../signature_test.png',
                    #     'outline': [0.2, 0.3, 0.5],
                    #     'border': 2,
                    #     'labels': False,
                    #     'display': 'CN,contact,date'.split(','),
                    #     # 'display': signature,
                    #     },

                    "signature_manual": [
                        ['background', 1, 0, 1],
                        # ['fill_colour', 1, 0, 0],
                        # ['rect_fill', 0, 50, 108, 150], ,['outline',1, 1, 0.5],
                        ['stroke_color', 0, 0, 1],  # border color
                        ['border', 1],
                        ['fill_colour', 0, 0, 1],  # text color
                        ['text_box', f'Документ подписан электронно', 'default', 0, 60, 270, 10, 7, True, 'center',
                         'top'],
                        ['text_box', f'Подписал(ФИО):\nEmail:\nДолжность:\nВремя подписания\nСертификат:',

                         # font  *[bounding box], size, wrap, align, baseline, spacing
                         'default', 5, 15, 80, 40, 7, True, 'left', 'top'],

                        ['text_box', f'{name}\n {names["E"]}\n{func_position}\n{signature_date}\n{fingerprint}',
                         'ttf_font', 90, 15, 270, 40, 7, True, 'left', 'top'],

                    ],

                    "contact": names['E'],
                    "location": "Russia",
                    "signingdate": date,
                    "reason": "Подписание",
                    # "signature": signature,
                    "password": password,

                }

            print(f'{name}\n{names["E"]}\n{func_position}\n{fingerprint}')
            # sign page with cryptography
            datau = open(input_file, 'rb').read()
            datas = pdf.cms.sign(datau,
                                 dct,
                                 p12pk, p12pc, p12oc,
                                 'sha1'
                                 )

            output_file = input_file.replace(input_file, dest)
            with open(output_file, 'wb') as fp:
                fp.write(datau)
                fp.write(datas)
            first_time = False
        return True
    except Exception as e:
        print(e)
        return False


if __name__ == '__main__':
    input_file = '1.pdf'
    pfx_certificate = 'container.pfx'
    password = 'password'
    dest = 'temp2.pdf'
    func_pos = 'TEST TEST'
    print(sign(func_pos,pfx_certificate,password,input_file,dest))
