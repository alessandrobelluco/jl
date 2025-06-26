import streamlit as st
st.set_page_config(layout='wide')
import pandas as pd
from PyPDF2 import PdfReader
from io import BytesIO
import xlsxwriter



st.title('POC - Importazione ordini')

path_pdf = st.sidebar.file_uploader('Caricare ordine in PDF')
if not path_pdf:
    #st.sidebar.warning('PDF ore non caricato')
    st.stop()


reader = PdfReader(path_pdf)
pages = reader.pages
#st.write(pages[1])

words = []

for page in pages:
        text = page.extract_text()
        words_i = list(text.split())
        words += words_i
        #words.append(words_i)

#Elaborazione====================================================================================================================


# il primo marker per identificare la riga è P.zo, i successivi  "---------"

db = pd.DataFrame(columns=['Indice','Tipo','Posizione','Codice','Descrizione','Qty','Disegno','Finitura','Altezza','Larghezza','Data_consegna'])



for i in range(len(words)-1):
    check = words[i]

    # condizioni per la prima posizione dell'ordine
    if (check == 'P.zo') and (words[i+1] != 'Pag.'): #TIPO1
        
        pos = words[i+1]
        codice = words[i+2]
        descrizione = words[i+3]
        #identifico il placeholder "PZ" che serve per indicare la fine della descrizione
        k=1
        while words[i+3+k] != 'PZ':
             descrizione += f' {words[i+3+k]}'
             k+=1
        placeh = k+i+3
        #placeh
        qty = words[placeh+1]
        datacons = words[placeh+2]
        #identifico il dis
        for a in range(placeh, len(words)):
            if words[a]=='Dis.':
                  disegno=words[a+1]

            if words[a]=='Altezza':
                alt = words[a+2]
                um_alt = words[a+3]
            if words[a]=='Larghezza':
                lar = words[a+2]
                um_lar = words[a+3]
            if words[a]=='Finitura':
                fin = words[a+1]
                break

        # nuova riga del db
        db.loc[len(db)]=[None,None,None,None,None,None,None,None,None,None,None]
        db.Posizione.iloc[-1]=pos
        db.Codice.iloc[-1]=codice
        db.Descrizione.iloc[-1]=descrizione
        db.Qty.iloc[-1]=qty
        db.Disegno.iloc[-1]=disegno
        db.Finitura.iloc[-1]=fin
        db.Altezza.iloc[-1]=alt
        db.Larghezza.iloc[-1]=lar
        db.Data_consegna.iloc[-1]=datacons
        db.Tipo.iloc[-1]='Tipo1'

    elif (check == 'P.zo') and (words[i+1] == 'Pag.'): #TIPO2
         #i
         #non dovrebbe fare niente
         pass

    elif (check == '---------') and words[i+2] != 'Valore' and words[i+1] != 'Legenda' : #TIPO3
        pos = words[i+1]
        codice = words[i+2]
        descrizione = words[i+3]
        k=1
        while (words[i+3+k] != 'PZ') and (words[i+3+k] != 'CM2') and (words[i+3+k][-2:] != 'PZ') :
            descrizione += f' {words[i+3+k]}'
            k+=1
            if k+i+3 == len(words):
                break
            
        if k+i+3 != len(words): #per risolvere il problema del placeholder in fondo alla lista senza effettive posizioni successive

            placeh = k+i+3

            try:
                qty = words[placeh+1]
            except:
                placeh

            datacons = words[placeh+2]
            
            for a in range(placeh, len(words)):
                if words[a]=='Dis.':
                    disegno=words[a+1]

                if words[a]=='Altezza':
                    alt = words[a+2]
                    um_alt = words[a+3]
                if words[a]=='Larghezza':
                    lar = words[a+2]
                    um_lar = words[a+3]
                if words[a]=='Finitura':
                    fin = words[a+1]
                    break

            # nuova riga del db
            db.loc[len(db)]=[None,None,None,None,None,None,None,None,None,None,None]
            db.Indice.iloc[-1]=i
            db.Posizione.iloc[-1]=pos
            db.Codice.iloc[-1]=codice
            db.Descrizione.iloc[-1]=descrizione
            db.Qty.iloc[-1]=qty
            db.Disegno.iloc[-1]=disegno
            db.Finitura.iloc[-1]=fin
            db.Altezza.iloc[-1]=alt
            db.Larghezza.iloc[-1]=lar
            db.Data_consegna.iloc[-1]=datacons
            db.Tipo.iloc[-1]='Tipo3'
 
    elif (check == '---------') and words[i+2] != 'Valore' and words[i+1] == 'Legenda' : #TIPO4
        k=0
        while words[i+k] != 'acquisto':
            k+=1
        place=k
        pos = words[i+k+5]
        codice = words[i+k+6]
        descrizione = words[i+k+7]
        n=1
        while (words[i+7+k+n] != 'PZ') and (words[i+7+k+n] != 'CM2') and (words[i+7+k+n][-2:] != 'PZ') :
            descrizione += f' {words[i+7+k+n]}'
            n+=1
        placeh = k+i+7+n
        qty = words[placeh+1]
        datacons = words[placeh+2]

        for a in range(placeh, len(words)):
            if words[a]=='Dis.':
                  disegno=words[a+1]

            if words[a]=='Altezza':
                alt = words[a+2]
                um_alt = words[a+3]
            if words[a]=='Larghezza':
                lar = words[a+2]
                um_lar = words[a+3]
            if words[a]=='Finitura':
                fin = words[a+1]
                break
            # nuova riga del db
        db.loc[len(db)]=[None,None,None,None,None,None,None,None,None,None,None]
        db.Indice.iloc[-1]=i
        db.Posizione.iloc[-1]=pos
        db.Codice.iloc[-1]=codice
        db.Descrizione.iloc[-1]=descrizione
        db.Qty.iloc[-1]=qty
        db.Disegno.iloc[-1]=disegno
        db.Finitura.iloc[-1]=fin
        db.Altezza.iloc[-1]=alt
        db.Larghezza.iloc[-1]=lar
        db.Data_consegna.iloc[-1]=datacons
        db.Tipo.iloc[-1]='Tipo4'


speciali = [
    'SO-OBLIQ.',
    'SV-OBLIQ.',
    'ZED',
    'DIP',
    'DMRDIP',
    'STREAM',
]

#db['Lavorazioni'] = ['ZED' in check for check in db.Descrizione]
db.drop(columns=['Indice','Tipo'], inplace=True)

db['Lavorazioni'] = None

for i in range(len(db)):
    check_des = db.Descrizione.iloc[i]
    check_disegno = db.Disegno.iloc[i]
    for lav in speciali:
        if (lav in check_des) or (lav in check_disegno):
            db['Lavorazioni'].iloc[i] = lav
        else:
            db['Lavorazioni'].iloc[i] = '-'


def scarica_excel(df, filename):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1',index=False)
    writer.close()

    st.download_button(
        label="Download Excel workbook",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.ms-excel"
    )


db
scarica_excel(db, 'Ordine_elaborato.xlsx')
