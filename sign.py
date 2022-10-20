
import argparse
import PyPDF2
import re
import sys
import datetime
from cryptography.hazmat import backends
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.x509.oid import NameOID
from endesive import pdf
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

def sign(pfx_certificate, password,input_file, dest):
    try:
        p12pk, p12pc, p12oc = load_pfx(pfx_certificate, password)

        names = get_rdns_names(p12pc.subject.rdns)
        # pdf_f = open(input_file,'rb')
        # pdfread = PyPDF2.PdfFileReader(pdf_f)
        # # num_pages = pdfread.getNumPages()
        dest = dest if dest else input_file


        current_date = datetime.datetime.now()
        date = current_date.strftime(f'%Y%m%d%H%M%S+03\'00\'')
        beautiful_date = current_date.strftime('%d/%m/%Y %H:%M:%S')
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
            mask = f"Signature"
            if mask in sign:
                i+=1
        # print(i)


        # offset coordinates of the digital sign
        if i == 1:
            plus_coord = 0
        else:
            plus_coord = 130 * (i-1)

        # parameters of the sign like size, align, place coords, etc
        # signature = signature_string(names['CN'], date, names['C'])
        dct = {
        "aligned": 0,
        "sigflags": 3,
        "sigflagsft": 132,
        "sigpage": 0,
        'sigandcertify': True,
        "auto_sigfield": False,
        #"sigandcertify": False,
        "signaturebox": (20 + plus_coord, 20, 130 + plus_coord, 75),
        "signform": False,
        "sigfield": f"Signature{i}",
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
            ['text_box', f'{names["CN"]}\n{names["E"]}\n{beautiful_date}',
             # font  *[bounding box], size, wrap, align, baseline, spacing
             'default', 5, 10, 270, 40, 7, True, 'left', 'top'],
            ['fill_colour', 0.4, 0.4, 0.4],
            ['rect_fill', 1, 50, 108, 1],
            ['fill_colour', 0, 0, 0],['border', 1],['outline',0.2, 0.3, 0.5],['background',1,1,1],
            ['stroke_color',0.5,1,1]

        ],

        "contact": names['E'],
        "location": "Russia",
        "signingdate": date,
        "reason": "Подписание",
        # "signature": signature,
        "password": password,

        }


        # sign page with cryptography
        datau = open(input_file, 'rb').read()
        datas = pdf.cms.sign(datau,
                     dct,
                     p12pk, p12pc, p12oc,'sha256'
                     )


        output_file = input_file.replace(input_file, dest)
        with open(output_file, 'wb') as fp:
          fp.write(datau)
          fp.write(datas)
        pdf_f.close()
        return True
    except Exception as e:
        print(e)
        return False


if __name__ == '__main__':
    input_file = 'temp1.pdf'
    pfx_certificate = 'container.pfx'
    password = '123'
    dest = 'temp2.pdf'

    print(sign(pfx_certificate,password,input_file,dest))
