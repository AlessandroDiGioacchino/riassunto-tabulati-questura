
# import openpyxl  # Needed for pd.DataFrame.to_excel
import os
import pandas as pd
import time

from PyPDF2 import PageObject, PdfReader
from collections import defaultdict


def main():
    
    path = input( 'Trascina qui il PDF e premi `invio`: ' )
    if path.startswith( '"' ):
        path = path[ 1: ]
        path = path[ : -1 ]

    path = os.path.normpath( path )
    reader = PdfReader( path )
    
    info = defaultdict( list )

    for page in reader.pages:
        for key, value in scrape_info( page ).items():
            info[ key ].append( value )

    output = pd.DataFrame.from_dict( info )
    index = pd.Index( range( 1, len( reader.pages ) + 1 ) )
    output.sort_values( by='Sezione', inplace=True )
    output.set_index( [ index, 'id' ], inplace=True )

    # current_datetime = datetime.datetime.now().strftime( '%Y-%m-%d %H%M%S' )
    filename = os.path.basename( path ).split( '.' )[ 0 ] + '.xlsx'
    output.to_excel( os.path.dirname( path ) + '\\' + filename )

    print( f'Ho scritto i risultati in {os.path.dirname( path )}\\' + filename )
    time.sleep( 4 )
          
    over = input( 'Premi `invio` per uscire dal programma' )

    while over != '':
        time.sleep( 0.5 )


def scrape_info( page: PageObject ) -> dict:

    id_, name, surname = '', '', ''
    section = ''
    place, start, end = '', '', ''
    gross, deposit, net = '', '', ''

    words = page.extract_text().split()
    words = list( map( lambda w: w.lower(), words ) )

    for ix, word in enumerate( words ):
        if word == 'sezione' and words[ ix + 1 ].isdecimal():
            section = words[ ix + 1 ]

        if word == 'id.' and words[ ix + 1 ] == 'dipendente':
            id_ = words[ ix + 2 ]

        if word == 'cognome':
            surname_ = words[ ix + 1 ]
        
            ix_ = ix + 2
            while words[ ix_ ] != 'nome':
                surname_ = surname_ + ' ' + words[ ix_ ]
                ix_ = ix_ + 1
            
            surname = surname_

        if word == 'nome':
            name_ = words[ ix + 1 ]
        
            ix_ = ix + 2
            while words[ ix_ ] != 'nato' and words [ ix_ + 1 ] != 'il':
                name_ = name_ + ' ' + words[ ix_ ]
                ix_ = ix_ + 1
        
            name = name_

        if word == "localita'" and words[ ix + 1 ] == 'missione':
            place_ = words[ ix + 2 ]

            ix_ = ix + 3
            while words[ ix_ ] != 'periodo' and words[ ix_ + 2 ] != 'missione':
                place_ = place_ + ' ' + words[ ix_ ]
                ix_ = ix_ + 1

            place = place_

        if word == 'periodo' and words[ ix + 2 ] == 'missione:':
            start = words[ ix + 3 ]
            end = words[ ix + 4 ]

        if word == 'totale' and words[ ix + 1 ] == 'dovuto':
            gross = words[ ix + 4 ]

        if word == 'acconto' and words[ ix + 1 ] == 'ricevuto' \
            and words[ ix + 2 ] == 'dalla':
        
            if words[ ix + 6 ] == 'importo':
                deposit = '0,00'
            else:
                deposit = words[ ix + 6 ]

        if word == 'importo' and words[ ix + 2 ] == 'mandato':
            net = words[ ix + 8 ]

    
    info = { 'id': id_, 'Nome': name, 'Cognome': surname, 'Sezione': section,
             'Località': place, 'inizio': start, 'Fine': end, 'Lordo': gross,
             'Acconto': deposit, 'Netto': net }

    return info


if __name__ == '__main__':
    main()
