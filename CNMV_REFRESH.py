import os
import io
import pandas as pd
import numpy as np

from pandas.errors import EmptyDataError
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation 
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.utils import quote_sheetname

from unicodedata import normalize
from datetime import datetime, timedelta, date
import calendar 
 
from bs4 import BeautifulSoup
from urllib.request import urlopen
import urllib.request
import urllib 

def add90days(date, n_days):
    if type(date) == np.str:
        date = datetime.strptime(date, '%d/%m/%Y')
        new_date = date + timedelta(days=n_days)
        return new_date
    else:
        new_date = date + timedelta(days=n_days)
        return new_date

def lastmonthday(date):
    if type(date) == np.str:
        date = datetime.strptime(date, '%d/%m/%Y')
        new_date = datetime(date.year, date.month, calendar.monthrange(date.year,date.month)[1])
    else:
        new_date = datetime(date.year, date.month, calendar.monthrange(date.year,date.month)[1])
    return new_date

def make_soup(url):	
    html = urlopen(url).read()
    return BeautifulSoup(html, 'lxml')

cnmv_folder = os.getcwd()
sociedades_folder = cnmv_folder + '/CNMV_Sociedades_TEST'
fondos_folder = cnmv_folder+'/CNMV_Fondos_TEST'
directorio_folder = cnmv_folder+'/DIRECTORIOS'
test_folder = cnmv_folder+'/CNMV_Download_TEST'
#gestoras_folders = [x.split('_')[0] for x in os.listdir(sociedades_folder) if '.csv' not in x]

os.chdir(cnmv_folder)
listado = pd.read_excel('Listado SGIICs.xlsx')
cifs = listado['CIF'].to_list()
listado.drop('CIF', axis = 1, inplace = True)
cifs2 = []
for item in cifs:
    cifs2.append(normalize('NFKD',item).strip(' '))
listado['CIF'] = cifs2  

os.chdir(directorio_folder)
directorio = pd.read_excel('Directorio_CNMV_05May2021.xlsx')
os.chdir(cnmv_folder)

directorio3 = pd.DataFrame(columns = ['Gestora','Tipo_Vehiculo','NIF'])
os.chdir(sociedades_folder)
for item in [x.split("_")[0] for x in os.listdir(sociedades_folder) if '.csv' not in x]:
    os.chdir(sociedades_folder)
    try:
        os.chdir(item+'_SOCIEDADES')
    except:
        continue
    for item2 in os.listdir():
        directorio3.loc[len(directorio3.index)] = [item, 'SOCIEDAD',item2.split('.')[0]]
        
os.chdir(fondos_folder)
for item in [x.split("_")[0] for x in os.listdir(fondos_folder) if '.csv' not in x]:
    os.chdir(fondos_folder)
    try:
        os.chdir(item+'_FONDOS')
    except:
        continue
    for item2 in os.listdir():
        directorio3.loc[len(directorio3.index)] = [item, 'FONDO',item2.split('.')[0]]

checking = pd.DataFrame(columns = ['GESTORA', 'CIF_GESTORA', 'TIPO', 'VEHICULO', 'NIF_VEHICULO','ONLINE?','DOWNLOADED?'])
for item in directorio['NIF_VEHICULO'].to_list():
    if item not in directorio3['NIF'].to_list():
        #display(directorio[directorio['NIF_VEHICULO'] == item])
        #checking = checking.append(directorio[directorio['NIF_VEHICULO'] == item])
        row = directorio[directorio['NIF_VEHICULO'] == item]
        checking.loc[len(checking.index)] = [row['GESTORA'].iloc[0],row['CIF_GESTORA'].iloc[0], row['TIPO'].iloc[0], row['VEHICULO'].iloc[0], row['NIF_VEHICULO'].iloc[0], 'YES','NO']
    else:
        row = directorio[directorio['NIF_VEHICULO'] == item]
        checking.loc[len(checking.index)] = [row['GESTORA'].iloc[0],row['CIF_GESTORA'].iloc[0], row['TIPO'].iloc[0], row['VEHICULO'].iloc[0], row['NIF_VEHICULO'].iloc[0], 'YES','YES']

os.chdir(cnmv_folder)

trim1 = [datetime.strftime(datetime(year = x, month = 3, day = 31), '%d/%m/%Y') for x in range(2001,2023)]
sem1 = [datetime.strftime(datetime(year = x, month = 6, day = 30), '%d/%m/%Y') for x in range(2001,2023)]
trim3 = [datetime.strftime(datetime(year = x, month = 9, day = 30), '%d/%m/%Y') for x in range(2001,2023)]
sem2 = [datetime.strftime(datetime(year = x, month = 12, day = 31), '%d/%m/%Y') for x in range(2001,2023)]

dates_dict2 = {'Trimestre 1':'31/3/',
                'Semestre 1':'30/6/',
                'Trimestre 3':'30/9/',
                'Semestre 2':'31/12/'}

CIF_lists = ['Listado EGFPs.xlsx', 'Listado SGIICs.xlsx']

#egfps = pd.read_excel(CIF_lists[0])
sgiics = pd.read_excel(CIF_lists[1])

fecha = ['31/03/{}'.format(x) for x in range(2001, 2022)]
fecha.reverse()

refresh_CIFS = checking[checking['DOWNLOADED?'] == 'NO']
print("There are {} FONDOS to download and {} SOCIEDADES to download.".format(refresh_CIFS[refresh_CIFS['TIPO'] == 'FONDO'].shape[0], refresh_CIFS[refresh_CIFS['TIPO'] == 'SOCIEDAD'].shape[0]))
#refresh_CIFS

refresh_list = pd.DataFrame(columns = ['Institution','CIF'])
for item in refresh_CIFS:
    refresh_list = refresh_list.append(listado[listado['CIF'] == item])
                                               
today = date.today().strftime('%d/%m/%Y')
errors = []

all_dataframes = []
final_dataframe = pd.DataFrame()

data_count = 0
folders = [sociedades_folder, fondos_folder]
for folder in folders:
    os.chdir(folder)
    for folder2 in os.listdir(folder):
        if '.csv' not in folder2:
            os.chdir(folder2)
            data_count += len(os.listdir())
            os.chdir(folder)

for cif_indie in refresh_CIFS.index:
    data_count +=1
    os.chdir(cnmv_folder)
    cif = refresh_CIFS['CIF_GESTORA'].iloc[cif_indie]
    gestora_gestora = refresh_CIFS['GESTORA'].iloc[cif_indie]
    print('Current Fund Manager: {} - {} \nCurrent progress: {} out of {} fund managers. ({}%)'.format(refresh_CIFS['CIF_GESTORA'].iloc[cif_indie], refresh_CIFS['GESTORA'].iloc[cif_indie], data_count, directorio.shape[0], ((data_count/directorio.shape[0])*100)))
    
    if refresh_CIFS.iloc[cif_indie]['TIPO'] == 'FONDO':
        os.chdir(fondos_folder)
        dir = os.path.join(os.getcwd(), refresh_CIFS['GESTORA'].iloc[cif_indie]+"_FONDOS")
        link_gestora = 'https://www.cnmv.es/Portal/Consultas/IIC/SGIIC.aspx?nif='+cif+'&vista=5&fs='+today
        try:
            pagina_nif = make_soup(link_gestora)
        except:
            errors.append(link_gestora)
            print(errors)
            
    elif refresh_CIFS.iloc[cif_indie]['TIPO'] == 'SOCIEDAD':
        os.chdir(sociedades_folder)
        dir = os.path.join(os.getcwd(), refresh_CIFS['GESTORA'].iloc[cif_indie]+"_SOCIEDADES")
        link_gestora = 'https://www.cnmv.es/Portal/Consultas/IIC/SGIIC.aspx?nif='+cif+'&vista=6&fs='+today
        try:
            pagina_nif = make_soup(link_gestora)
        except:
            errors.append(link_gestora)
            print(errors)
    else:
        print('Something going on???')
    
    dir1 = os.path.join(fondos_folder, refresh_CIFS['GESTORA'].iloc[cif_indie]+"_FONDOS")
    dir2 = os.path.join(sociedades_folder, refresh_CIFS['GESTORA'].iloc[cif_indie]+"_SOCIEDADES")

    if not os.path.exists(dir1):
        os.mkdir(dir1)
    if not os.path.exists(dir2):
        os.mkdir(dir2)

    #if not os.path.exists(dir):
    #    os.mkdir(dir)
    os.chdir(dir)
    
    nifs = {}
    title = refresh_CIFS.iloc[cif_indie]['VEHICULO']
    NIF = refresh_CIFS.iloc[cif_indie]['NIF_VEHICULO']
    nifs[title] = NIF
    
    #Fechas aqui 

    main_link = "https://www.cnmv.es/Portal/Consultas"
    main_link_pdfs = "https://www.cnmv.es/Portal"
    full_links = []
    xbrl_data = []
    #all_dataframes = []
    currency_codes = ['', 'aed', 'afn', 'all', 'amd', 'ang', 'aoa', 'ars', 'aud', 'awg', 'azn', 'bam', 'bbd', 'bdt', 'bgn', 'bhd', 'bif', 'bmd', 'bnd', 'bob','bov', 'brl', 'bsd', 'btn', 'bwp', 'byn', 'bzd', 'cad', 'cdf', 'che', 'chf', 'chw', 'clf', 'clp', 'cny', 'cop', 'cou', 'crc', 'cuc', 'cup', 'cve', 'czk', 'djf', 'dkk', 'dop', 'dzd', 'egp', 'ern', 'etb', 'eur', 'fjd', 'fkp', 'gbp', 'gel', 'ghs', 'gip', 'gmd', 'gnf', 'gtq', 'gyd', 'hkd', 'hnl', 'hrk', 'htg', 'huf', 'idr', 'ils', 'inr', 'iqd', 'irr', 'isk', 'jmd', 'jod', 'jpy', 'kes', 'kgs', 'khr', 'kmf', 'kpw', 'krw', 'kwd', 'kyd', 'kzt', 'lak', 'lbp', 'lkr', 'lrd', 'lsl', 'lyd', 'mad', 'mdl', 'mga', 'mkd', 'mmk', 'mnt', 'mop', 'mru', 'mur', 'mvr', 'mwk', 'mxn', 'mxv', 'myr', 'mzn', 'nad', 'ngn', 'nio', 'nok', 'npr', 'nzd', 'omr', 'pab', 'pen', 'pgk', 'php', 'pkr', 'pln', 'pyg', 'qar', 'ron', 'rsd', 'rub', 'rwf', 'sar', 'sbd', 'scr', 'sdg', 'sek', 'sgd', 'shp', 'sll', 'sos', 'srd', 'ssp', 'stn', 'svc', 'syp', 'szl', 'thb', 'tjs', 'tmt', 'tnd', 'top', 'try', 'ttd', 'twd', 'tzs', 'uah', 'ugx', 'usd', 'usn', 'uyi', 'uyu', 'uzs', 'vef', 'vnd', 'vuv', 'wst', 'xaf', 'xcd', 'xdr', 'xof', 'xpf', 'xsu', 'xua', 'yer', 'zar','zmw','zwl']
    #currency_codes = ['eur','gbp','usd','hkd','aud','cad','chf','dkk','jpy','nok','nzd','sek','sgd','brl','mxn','cop','clp','pen','zar','krw','thb','pln','twd','php','try','myr','idr','egp','huf',]
    tagname_dict = {'MAIN':'iic-com:inversionesfinancierasexterior',
                    'INDIVIDUAL':'iic-com:inversionesfinancierasrvcotizada',
                    'ISIN':'iic-com:codigoisin',
                    'NAME':'iic-com:inversionesfinancierasdescripcion',
                    'CURRENCY':'dgi-lc-int:xcode_iso4217.',
                    'CURRENCY2':'xcode_iso4217.',
                    'VALUE':'iic-com:inversionesfinancierasvalor'}

    nif_index = 1

    for codigo_nif in list(nifs.values()):
        counting = 0
        tipo_tipo = refresh_CIFS.iloc[cif_indie]['TIPO']
        if tipo_tipo == 'FONDO':
            location = fondos_folder + '/' + gestora_gestora + '_FONDOS'
        elif tipo_tipo == 'SOCIEDAD':
            location = sociedades_folder + '/' + gestora_gestora + '_SOCIEDADES'
        nifs_downloaded = len(os.listdir(location))
        nifs_total = directorio[(directorio['GESTORA'] == gestora_gestora) & (directorio['TIPO'] == tipo_tipo)].shape[0]
            
        print("Currently downloading data for {}: {} \nThe fund is {} out of {}".format(refresh_CIFS.iloc[cif_indie]['TIPO'], codigo_nif, nifs_downloaded, nifs_total))
        nif_dataframe = pd.DataFrame()
        for date in fecha:
            if refresh_CIFS.iloc[cif_indie]['TIPO'] == 'FONDO':
                link_nif = "https://www.cnmv.es/Portal/Consultas/IIC/Fondo.aspx?nif="+ codigo_nif +"&vista=1&fs=" + date
            elif refresh_CIFS.iloc[cif_indie]['TIPO'] == 'SOCIEDAD':
                link_nif = "https://www.cnmv.es/Portal/Consultas/IIC/SociedadIIC.aspx?nif="+ codigo_nif +"&vista=1&fs=" + date
            else:
                print('wtf?!?!')
                
            pagina_individual = make_soup(link_nif)
            
            period = [x.contents[0].strip() for x in pagina_individual.find_all('td') if x.attrs['data-th'] == 'Periodo']
            year = [x.contents[0] for x in pagina_individual.find_all('td') if x.attrs['data-th'] == 'Ejercicio']
            periodo2 = [[period[x],year[x]] for x in range(len(period))]

            #periodo = dict(zip([x.contents[0].strip() for x in pagina_individual.find_all('td') if x.attrs['data-th'] == 'Periodo'], [x.contents[0] for x in pagina_individual.find_all('td') if x.attrs['data-th'] == 'Ejercicio']))
            #print(periodo2)
            if len(periodo2) == 0:
                break

            xbrl_links = [main_link+x.find_all('a')[1]['href'].strip('.') for x in pagina_individual.find_all('td') if x.attrs['data-th'] == 'Documentos']
            pdf_links = [main_link_pdfs+'/'+x.find_all('a')[0]['href'].strip('/..') for x in pagina_individual.find_all('td') if x.attrs['data-th'] == 'Documentos']
            
            #print(link_nif)
            #print(pdf_links)
            
            for link_indie in range(len(xbrl_links)):
                #print("Periodo: {} para el Año: {}".format(list(periodo.keys())[link_indie], list(periodo.values())[link_indie]))
                print("Periodo: {} para el Año: {}".format(periodo2[link_indie][0], periodo2[link_indie][1]))
                #fechacha = dates_dict2[list(periodo.keys())[link_indie]] + list(periodo.values())[link_indie]
                fechacha = dates_dict2[periodo2[link_indie][0]] + periodo2[link_indie][1]
                full1 = make_soup(xbrl_links[link_indie])
                data = full1.find_all(tagname_dict['MAIN'])
                
                nombre_vehiculo = []
                nombre_vehiculo2 = []
                for item in full1.find_all('iic-com-fon:denominacionfondo'):
                    nombre_vehiculo.append(item.text)
                for item in full1.find_all('denominacionsociedad'):
                    nombre_vehiculo.append(item.text)
                for item in full1.find_all('identifier'):
                    nombre_vehiculo.append(item.text)

                if not nombre_vehiculo:
                    for item in full1.find_all('xbrli:identifier'):
                        nombre_vehiculo2.append(item.text)
                    try:
                        nombre_ind = nombre_vehiculo2[0]
                    except:
                        print("The following fund has no name: {} \nCheck here: {}".format(title, link_nif))
                else:
                    nombre_ind = nombre_vehiculo[0]

                isins = []
                names = []
                currency = []
                values = []
                for item in data:
                    individuals = item.find_all(tagname_dict['INDIVIDUAL'])
                    for item2 in individuals:
                        data_isins = item2.find_all(tagname_dict['ISIN'])
                        data_names = item2.find_all(tagname_dict['NAME'])
                        data_currency = item2.find_all([tagname_dict['CURRENCY']+x for x in currency_codes] + [tagname_dict['CURRENCY2']+x for x in currency_codes])
                        for isin in data_isins:
                            isins.append(str(isin.text))
                        for name in data_names:
                            names.append(str(name.text))
                        for cur in data_currency:
                            currency.append(str(cur.text))

                        data_values = item2.find_all(tagname_dict['VALUE'])
                        if len(data_values) == 2:
                            #print([x.text for x in data_values if 'ia' in x['contextref']][0])
                            #count += 1
                            vals = [x.text for x in data_values if 'ia' in x['contextref']]
                            if len(vals) > 0:
                                values.append([x.text for x in data_values if 'ia' in x['contextref']][0])
                            else:
                                print(data_values)
                        elif len(data_values) == 1 and 'ia' in data_values[0]['contextref']:
                            #print(data_values[0].text)
                            #count +=1
                            values.append(data_values[0].text)
                        elif len(data_values) == 1 and ('ipp' in data_values[0]['contextref'] or 'ipy' in data_values[0]['contextref']):
                            #print('NO CURRENT VALUE')
                            #count +=1
                            values.append(0)
                        elif len(data_values) == 0:
                            pass

                dataff = pd.DataFrame()
                dataff['NOMBRE_VEHICULO'] = [nombre_ind for x in isins]
                dataff['VEHICULO'] = [codigo_nif for x in isins]
                dataff['DATE'] = [fechacha for x in isins]
                dataff['ISIN'] = isins
                dataff['NAME'] = names
                try:
                    dataff['CURRENCY'] = currency
                except ValueError:
                    print('{}. The currency data is mismatched from the following date: {}. \nXBRL Link: {} \nPDF Link: {}'.format(counting, date, xbrl_links[link_indie],pdf_links[link_indie]))
                    print(link_nif)
                try:
                    dataff['VALUE'] = values
                except ValueError:
                    print('{}. The value data from {} could not be parsed cause of a mismatched list length. \nXBRL Link: {}\nPDF Link: {}'.format(counting, date, xbrl_links[link_indie], pdf_links[link_indie]))
                    print(link_nif)
                    print(values)
                    break
                nif_dataframe = nif_dataframe.append(dataff.drop_duplicates())
                counting+=1

        nif_dataframe.reset_index(drop = True, inplace = True)
        #nif_dataframe.to_csv("{}_{}.csv".format(nif_index, fondo), index = False
        nif_dataframe.to_csv("{}.csv".format(codigo_nif), index = False)
        #all_dataframes.append(nif_dataframe.reset_index(drop = True))
        nif_index += 1
    #Standardizing Fund Name to most recent
    for item in all_dataframes:
        if item.empty:
            pass
        try:
            name_dict = {a:item[item['DATE'] == item['DATE'].max()]['NOMBRE_VEHICULO'].iloc[0] for a in list(item['NOMBRE_VEHICULO'].unique())}
            item['NOMBRE_VEHICULO'].replace(name_dict, inplace = True)
        except:
            print(item)
    
    os.chdir(cnmv_folder)

    manager_name = refresh_CIFS.iloc[cif_indie]['GESTORA']

    #final_dataframe = pd.DataFrame()
    #for item in all_dataframes:
    #    final_dataframe = final_dataframe.append(item)
    #final_dataframe.reset_index(drop = True, inplace = True)

#Checking to see if all fondo and sociedad NIFS have downloaded for one gestora before downloading final csv file is created
    hestora = refresh_CIFS.iloc[cif_indie]['GESTORA']
    cif_hestora = refresh_CIFS.iloc[cif_indie]['CIF_GESTORA']
    
    if refresh_CIFS.iloc[cif_indie]['TIPO'] == 'FONDO':
        cnmvlist = directorio[(directorio['CIF_GESTORA'] == cif_hestora) & (directorio['TIPO'] == 'FONDO')]['NIF_VEHICULO'].to_list()
        dirlist = [x.split('.')[0] for x in os.listdir(fondos_folder + '/' + hestora + '_FONDOS')]
        dirrie = fondos_folder + '/' + hestora + '_FONDOS'
        dirrie2 = fondos_folder
        os.chdir(fondos_folder)
    elif refresh_CIFS.iloc[cif_indie]['TIPO'] == 'SOCIEDAD':
        cnmvlist = directorio[(directorio['CIF_GESTORA'] == cif_hestora) & (directorio['TIPO'] == 'SOCIEDAD')]['NIF_VEHICULO'].to_list()
        dirlist = [x.split('.')[0] for x in os.listdir(sociedades_folder + '/' + hestora + '_SOCIEDADES')]
        dirrie = sociedades_folder + '/' + hestora + '_SOCIEDADES'
        dirrie2 = sociedades_folder
        os.chdir(sociedades_folder)
    else:
        print('Ok now hwhat?!?!?!?!?!?!?!?!?!?!?!?!?!?!?!')

    if all([(x in cnmvlist) for x in dirlist]) and all([(x in dirlist) for x in cnmvlist]):
        print("All {} data has been downloaded for {} with CIF {} in the folder {}".format(dirrie.split('_')[-1], hestora, cif_hestora, dirrie))
        os.chdir(dirrie)
        for niffie in dirlist:
            try:
                data_recover = pd.read_csv("{}.csv".format(niffie))
                all_dataframes.append(data_recover)
            except EmptyDataError:
                print('DataFrame is empty for NIF {}'.format(niffie))

        for item in all_dataframes:
            final_dataframe = final_dataframe.append(item)
        os.chdir(dirrie2)

        final_dataframe.reset_index(drop = True, inplace = True)
        final_dataframe.to_csv(manager_name+'_CNMV_Data_Download.csv', index = False)
        all_dataframes = []
        final_dataframe = pd.DataFrame()
    else:
        #print("Somethin aint right???")
        pass