# hoja['A{}'.format(i)] = "ok"



def crearXlsx(directory):
    wb = Workbook()
    wb.create_sheet("DATOS",0)
    wb.save(directory)

def defineHeaderXlsx(directory,colunmName):
    wb = load_workbook(directory)
    hoja = wb['DATOS']
    con = 1
    for i in colunmName:
       hoja.cell(row=1, column=con).value = i
       con +=1
    wb.save(directory)


    
# creacion
# colunmName = ['ID_MUESTREO', 'FECHA', 'ESTACION',
#               'RESPONSABLE', 'ENTIDAD', 'Directorio completo']
# crearXlsx(directory2)
# defineHeaderXlsx(directory2,colunmName)


# datos
#full_file_paths = explorar(directoryraiz)
#

# add
# addData(directory2,registros,full_file_paths,colunmName)

# sheet = crearXlsx(directory2)



# registro = ReadAllRow(directory3)


# addData(directory2,registro,full_file_paths,colunmName)


# leer
#directory3 ="C:\Prueba\Reading.xlsx"
#directory4 ="C:\Prueba\SAMP.xlsm"
# ReadXlsx(directory4)


'''
print("===============Nombres======================")
name_file = get_nameFilepaths("C:\Prueba")
for r in name_file:
    print(r)
    
    
print("===============Full path======================")
full_file_paths = get_filepaths("C:\Prueba")
for r in full_file_paths:
    print(r)
print("=================Carpetas====================")
carpetas = get_folderpaths("C:\Prueba")
for r in carpetas:
    print(r)
'''


# crar xlsx
# crearXlsx(directory)


# cargar
# loadXlsx(directory)


# escribir
# writeXlsx(directory2)


# leer
#dato = ReadXlsx(directory2)
# 3print(dato)





def get_filepaths(directory):
    suffix = ".xlsm"
    registro = []

    contenido = os.listdir(directory)

    # obtener archivos
    carpetas = [nombre for nombre in contenido if isdir(
        join(directory, nombre))]

    for r in carpetas:
        # Walk the tree.
        for root, directories, files in os.walk(directory+'\\'+r):
            for filename in files:
                if(filename.endswith(suffix)):
                    registro.append(ReadXlsx(directory+'\\'+r+'\\'+filename))

    # imprimir
    for r in registro:
        print(r[0].value, r[1].value, r[2].value, r[3].value, r[4].value)

    return registro





def get_filepaths2(directory):
    suffix = ".xlsm"
    file_paths = []

    contenido = os.listdir(directory)

    # obtener archivos
    carpetas = [nombre for nombre in contenido if isdir(
        join(directory, nombre))]

    for r in carpetas:
        # Walk the tree.
        for root, directories, files in os.walk(directory+'\\'+r):
            for filename in files:
                # Join the two strings in order to form the full filepath.
                filepath = os.path.join(root, filename)
                if(filename.endswith(suffix)):
                    file_paths.append(filepath)  # Add it to the list.

    return file_paths




def get_nameFilepaths(directory):
    contenido = os.listdir(directory)
    file_name = []
    # obtener archivos
    carpetas = [nombre for nombre in contenido if isdir(
        join(directory, nombre))]

    for r in carpetas:
        # obtener archivos
        return [nombre for nombre in contenido if isfile(join(directory, nombre))]


def get_folderpaths(directory):

    contenido = os.listdir(directory)

    # obtener archivos
    return [nombre for nombre in contenido if isdir(join(directory, nombre))]

