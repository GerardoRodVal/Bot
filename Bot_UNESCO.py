# encoding utf-8

import requests as rq
import wget
import os
import zipfile37

import pandas as pd
from selenium import webdriver
import platform


def ultima_version(SO):
    ultima = rq.get('https://chromedriver.storage.googleapis.com/LATEST_RELEASE').text
    files = {'Linux': 'chromedriver_linux64.zip',
             'Darwin': 'chromedriver_mac64.zip',
             'Windows': 'chromedriver_win32.zip'}

    wget.download('https://chromedriver.storage.googleapis.com/' + ultima + '/' + files[SO])

    try:
        zip = zipfile37.ZipFile(files[SO])
        zip.extractall()
        os.system('chmod +x chromedriver')
    except:
        pass

    return os.remove(files[SO])


def sistema_operativo():
    browser = "Sistema Operativo"
    sistema = platform.system()
    SO = "{}".format(sistema)
    print('Sistema operativo: ', SO)

    ultima_version(SO)

    if (SO == "Linux" or SO == 'Darwin'):
        browser = webdriver.Chrome('./chromedriver')

    elif(SO == "Windows"):
        browser = webdriver.Chrome('./chromedriver_Win.exe')

    return browser


def registro(formulario, postales, browser):
    for ind in range(len(formulario)):
        # ---------------------------------------- Variables General -------------------------------------------------
        Pais = formulario['País'].iloc[ind]
        Nom_escuela = formulario['Nombre de la Escuela'].iloc[ind]
        EMAIL = formulario['Email de la Escuela'].iloc[ind]
        DPOSTAL = formulario['Dirección postal'].iloc[ind]
        CPOSTAL = postales[ind]
        CIUDAD = formulario['Ciudad/Poblado'].iloc[ind]
        ESTADO = formulario['Estado'].iloc[ind]
        TEL =   formulario['Teléfono (LADA)'].iloc[ind]
        WEB = formulario['Página Web'].iloc[ind]
        FACEBOOK = formulario['Facebook'].iloc[ind]
        TWITTER = formulario['Twitter'].iloc[ind]
        N_ENSE = formulario['Nivel de enseñanza'].iloc[ind]
        CATEGORIA = formulario['Categoría de la escuela'].iloc[ind]
        UBICACION = formulario['Ubicación'].iloc[ind]
        # --------------------------------------- Variables de Director ----------------------------------------------
        TITULOD = formulario['Titulo del director'].iloc[ind]
        NOMBRED = formulario['Nombre del Director'].iloc[ind]
        APELLIDOD = formulario['Apellido del Director'].iloc[ind]
        EMAILD = formulario['Email del Director'].iloc[ind]
        # ------------------------------------------ Variables de Focal ----------------------------------------------
        TITULOF = formulario['Titulo del encargado'].iloc[ind]
        NOMBREF = formulario['Nombre del encargado'].iloc[ind]
        APELLIDOF = formulario['Apellido del encargado'].iloc[ind]
        cargosF = ['Director', 'Administrador', 'Docente', 'Miembro del consejo Escolar']
        CARGOF = formulario['Cargo que ocupa el encargado'].iloc[ind]
        EMAILF = formulario['Email del encargado'].iloc[ind]
        TIEMPOF = formulario['Disponibilidad de tiempo para la implementación y planeación de actividades'].iloc[ind]
        # --------------------------------------------- Constantes ---------------------------------------------------
        RECURSOSF = 'Si'
        DETALLES = 'Knotion representa un medio a través del cuál dedicamos tiempo, recursos humanos y recursos' \
                   ' económicos para la implementación y desarrollo de proyectos alineados a los ODS; ' \
                   'aprovecharemos esta inversión para cumplir los requerimientos para pertenecer a la redPEA.'

        EXPLICACION = formulario['Por favor, explica porque quieres ser parte de la redPEA UNESCO (100 palabras máximo).'].iloc[ind]
        ENTERO = 'Knotion'

        # ----------------------------------------- CONECTANDO AL DRIVER CHROME ------------------------------------------------
        browser.get('https://aspnet.unesco.org/es-es/Paginas/RequestToBeAMember.aspx')                                # URL de la pagina a controlar

        # -------------------------------------------------- CATEGORIA GENERAL -------------------------------------------------
        Nom_Pais = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_ctl01editableRegion')   # ID del elemento caja de texto
        Nom_Pais.send_keys(Pais)                                                                                      # enviando informacion al elemento

        escuela = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Name')
        escuela.send_keys(Nom_escuela)

        email = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_email')
        email.send_keys(EMAIL)

        Dpostal = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Street')
        Dpostal.send_keys(DPOSTAL)

        Cpostal = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_PostCode')
        Cpostal.send_keys(CPOSTAL)

        Ciudad = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Town')
        Ciudad.send_keys(CIUDAD)

        Estado = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Province')
        Estado.send_keys(ESTADO)

        Tel = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_PhoneNumber')
        Tel.send_keys(TEL)

        web = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Website')
        web.send_keys(WEB)

        facebook = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Facebook')
        facebook.send_keys(FACEBOOK)

        twitter = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Twitter')
        twitter.send_keys(TWITTER)

        N_ense = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_ctl03editableRegion')
        N_ense.send_keys(N_ENSE)

        categoria = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_ctl05editableRegion')
        categoria.send_keys(CATEGORIA)

        ubicacion = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_ctl07editableRegion')
        ubicacion.send_keys(UBICACION)

        # ----------------------------------------------- CATEGORIA RESPONSABLE/DIRECTOR ---------------------------------------
        if( TITULOD == 'Señora' ):
            TituloD = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_PrincipalCivility')
            TituloD.send_keys(TITULOD)
        elif(TITULOD == 'Señorita'):
            TituloD = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_PrincipalCivility')
            TituloD.send_keys(TITULOD)

        NombreD = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_PrincipalFirstName')
        NombreD.send_keys(NOMBRED)

        ApellidoD = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_PrincipalLastName')
        ApellidoD.send_keys(APELLIDOD)

        EmailD = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_PrincipalEmail')
        EmailD.send_keys(EMAILD)

        #-------------------------------------------------- CATEGORIA PUNTO FOCAL ----------------------------------------------
        if(TITULOF == 'Señora'):
            TituloF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_FocalPointCivility')
            TituloF.send_keys(TITULOF)
        elif(TITULOF == 'Señorita'):
            TituloF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_FocalPointCivility')
            TituloF.send_keys(TITULOF)

        NombreF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_FocalPointFirstName')
        NombreF.send_keys(NOMBREF)

        ApellidoF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_FocalPointLastName')
        ApellidoF.send_keys(APELLIDOF)

        if( CARGOF not in cargosF ):
            CargoF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_FocalPositionList')
            CargoF.send_keys('Otro (Especifique)')
            CargoF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_FocalPointPosition')
            CargoF.send_keys(CARGOF)
        else:
            CargoF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_FocalPositionList')
            CargoF.send_keys(CARGOF)

        EmailF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_FocalPointEmail')
        EmailF.send_keys(EMAILF)

        TiempoF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_TimeResources')
        TiempoF.send_keys(TIEMPOF)

        RecursosF = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Cooperation')
        RecursosF.send_keys(RECURSOSF)

        Detalles = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_CooperationDetails')
        Detalles.send_keys(DETALLES)

        Explicacion = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_reason')
        Explicacion.send_keys(EXPLICACION)

        Entero = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_HearList')
        Entero.send_keys('Otro (Especifique)')
        Entero = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_s_Hear')
        Entero.send_keys(ENTERO)
        # -------------------------------------------- CATEGORIA PARA ENVIAR INFORMACION ---------------------------------------
        #Enviar = browser.find_element_by_id('ctl00_ctl47_g_4d272641_e1a3_4b25_8021_733ef70ad505_ctl00_Button2')
        #Enviar.click()                                                                                               # enviando el formulario


def Main_bot():
# ------------------------------------------ LECTURA DE ENTRADA ----------------------------------------------
    formulario = pd.read_csv('Registro para la redPEA.csv', encoding='utf-8')
    formulario = formulario.drop(columns='Marca temporal')

# ----------------------------------------- CORRECCION DE RESPUESTAS -----------------------------------------
    postales = ['0'*(5-len(str(i)))+str(i) for i in formulario['Código Postal']]
    formulario['Teléfono (LADA)'] = formulario['Teléfono (LADA)'].astype(str)
    formulario['Twitter'] = formulario['Twitter'].fillna(' ')
    formulario['Facebook'] = formulario['Facebook'].fillna(' ')
    formulario['Estado'] = formulario['Estado'].fillna(' ')

# ---------------------------------------------   Temporal ----------------------------------
    formulario['Por favor, explica porque quieres ser parte de la redPEA UNESCO (100 palabras máximo).'] = \
        formulario['Por favor, explica porque quieres ser parte de la redPEA UNESCO (100 palabras máximo).'].fillna(' ')

    browser = sistema_operativo()
    registro(formulario, postales, browser)

Main_bot()
