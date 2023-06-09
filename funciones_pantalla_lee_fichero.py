from googlemaps import Client
from itertools import product, permutations
import time
import copy
import random
import folium
import webbrowser
import shutil
import xlrd

#PARAMETROS
max_minutos_montaje = 550 #
min_minutos_montaje = 350
n_clientes = 7 #
criterio_importes_grandes = 1500 #
criterio_montajes_grandes = 200 #
n_maximo_montajes = 3 #
filtro_importe_ruta = 2600 #
importe_bajo = 500 #
importe_alto = 800 #
parametro_cliente_final = 30.0 #minutos al centro
lejos_almacen = 99 #estar a 60 min del almacén es un cliente lejos

#direccion_almacen = "Calle doctor Ferrán 43 - 08120 - La Llagosta - Barcelona"
#Barcelona 
#latitud_almacen = 41.5165833
#longitud_almacen = 2.2038206

direccion_almacen = "Avenida de Castilla 27 - 28830 - San Fernando de Henares - Madrid"
#Madrid 
latitud_almacen = 40.454164
longitud_almacen = -3.4969303

def leer_fichero_clientes(archivo):
    wb = xlrd.open_workbook(archivo)
    
    hoja = wb.sheet_by_index(0)

    matriz = []
    for i in range(hoja.nrows):
        fila = []
        for j in range(hoja.ncols):
            fila.append(hoja.cell_value(i, j))
        matriz.append(fila)
        i += 1
    
    return matriz

diccionario_de_clases = {'S' : 4, 'A' : 3, 'B' : 2, 'N' : 1, 'T' : 0, 'X' : 0, '': 0, '$' : 4}

diccionario_de_categorias = {'A lasciare' : 'T', 'Nessuna' : 'T', 'Armadio angolare' : 'M', 'Soggiorno lineare' : 'M', 'Armadio lineare' : 'M', 'Camera' : 'M', 'Cameretta angolare' : 'M', 'Cameretta lineare' : 'M', 'Cucina lineare p.misure' : 'C', 'Cucina lineare no p.misure' : 'C', 'Bagno Sospeso' : 'M', 'Cucina angolare p.misure' : 'R', 'Cucina angolare no p.misure' : 'R','Cucina doppio angolo p.misure': 'R'}

diccionario_de_categorias_2 = {'T' : 0, 'M' : 1, 'E' : 2, 'C' : 3, 'R' : 4}

diccionario_horas_gold = {1 : 8, 2 : 9, 3 : 10, 4 : 11, 5 : 12, 6 : 13, 7 : 14, 8 : 15, 9 : 16,}
#---------------------------------------------------------------------------------------------------------------------------------

#CLIENTES
#archivo = 'C:/Users/letic/OneDrive/Escritorio/Jaimebigdata/entregasdiasposteriores0803.xls'

clientes_gold = []
horas_gold = []
matriz_cliente_sin_duplicados = []
latitud = []
longitud = []
direcciones = []
lista_de_clases = []
lista_categoria_montaje = []
listado_importes = []
lista_de_tiempos_de_montaje = []
lista_de_cliente = []

matriz_cliente = []

def datos_clientes(archivo,almacennn):
    global clientes_gold
    global matriz_cliente_sin_duplicados
    global horas_gold
    global latitud
    global longitud
    global direcciones
    global lista_de_clases
    global lista_categoria_montaje
    global listado_importes
    global lista_de_tiempos_de_montaje
    global lista_de_cliente
    global clientes_faltantes
    global matriz_cliente
    
    matriz_cliente = leer_fichero_clientes(archivo)
    
    cliente = []
    pedidos = []
    matriz_cliente_sin_duplicados = []

    for a in matriz_cliente:
        if a[0] not in cliente:
            cliente.append(a[0])
            pedidos.append(a[1])
            matriz_cliente_sin_duplicados.append(copy.deepcopy(a))
            if a[6] != 'CatConsegna':
                matriz_cliente_sin_duplicados[cliente.index(a[0])][6] = diccionario_de_categorias[a[6]]
        else:
            if a[1] not in pedidos:
                pedidos.append(a[1])
                matriz_cliente_sin_duplicados[cliente.index(a[0])][3] = matriz_cliente_sin_duplicados[cliente.index(a[0])][3] + a[3]
                matriz_cliente_sin_duplicados[cliente.index(a[0])][8] += a[8]
                matriz_cliente_sin_duplicados[cliente.index(a[0])][9] += a[9]
                
                if diccionario_de_clases[a[2]] >= diccionario_de_clases[matriz_cliente_sin_duplicados[cliente.index(a[0])][2]]:
                    matriz_cliente_sin_duplicados[cliente.index(a[0])][2] = a[2]

                if diccionario_de_categorias_2[diccionario_de_categorias[a[6]]] >= diccionario_de_categorias_2[matriz_cliente_sin_duplicados[cliente.index(a[0])][6]]:
                    matriz_cliente_sin_duplicados[cliente.index(a[0])][6] = diccionario_de_categorias[a[6]]

                if a[7] > matriz_cliente_sin_duplicados[cliente.index(a[0])][7]:
                    matriz_cliente_sin_duplicados[cliente.index(a[0])][7] = a[7]   
            else:
                matriz_cliente_sin_duplicados[cliente.index(a[0])][8] += a[8]

        if a[10] == "FABRI-KIT S.L.U." and diccionario_de_categorias_2[matriz_cliente_sin_duplicados[cliente.index(a[0])][6]] < 2:
            matriz_cliente_sin_duplicados[cliente.index(a[0])][6] = 'E'

    matriz_cliente_sin_duplicados = [a for a in matriz_cliente_sin_duplicados if a[12] == almacennn]

    clientes_gold = [i+1 for i,a in enumerate(matriz_cliente_sin_duplicados) if a[7] != 0]
    horas_gold = [diccionario_horas_gold[a[7]] for i,a in enumerate(matriz_cliente_sin_duplicados) if a[7] != 0]

    latitud = [a[5] for a in matriz_cliente_sin_duplicados]
    longitud = [a[4] for a in matriz_cliente_sin_duplicados]
    #latitud.pop(0)
    #longitud.pop(0)

    direcciones = [[a[4],a[5]] for a in matriz_cliente_sin_duplicados]
    #direcciones.pop(0)

    lista_de_clases = [a[2] for a in matriz_cliente_sin_duplicados]
    #lista_de_clases.pop(0)

    lista_categoria_montaje2 = [a[6] for a in matriz_cliente_sin_duplicados]
    #lista_categoria_montaje2.pop(0)
    #lista_categoria_montaje = [diccionario_de_categorias[a] for a in lista_categoria_montaje2]
    lista_categoria_montaje = [a for a in lista_categoria_montaje2]

    listado_importes = [round(a[8],1) for a in matriz_cliente_sin_duplicados]
    #listado_importes.pop(0)

    lista_de_tiempos_de_montaje = [a[3] for a in matriz_cliente_sin_duplicados]
    #lista_de_tiempos_de_montaje.pop(0)

    lista_de_cliente = [a + 1 for a in range(len(direcciones))]
    
    clientes_faltantes = copy.deepcopy(lista_de_cliente)
    
    return clientes_gold, matriz_cliente_sin_duplicados, horas_gold, latitud, longitud, direcciones, lista_de_clases, lista_categoria_montaje, listado_importes, lista_de_tiempos_de_montaje, lista_de_cliente

#------------------------------------------------------------------------------------------------------------------------------------

#clientes_gold = [104]
#horas_gold = [9.0]
#rutas_disponibles = [14,6,2]

clientes_faltantes = copy.deepcopy(lista_de_cliente)
clientes_listos = []
clases_clientes_faltantes = copy.deepcopy(lista_de_clases)
clases_clientes_listos = []
rutas_listas = []
rutas_guardadas = []

#gmaps= Client(key="AIzaSyCfldJIo4Ffh3WFNqBR5VTzOA6F1kNyPf8")
#gmaps= Client(key="AIzaSyDkQJ7-L4VtxaAAO3BC4HVqES47vk4ssxs")

total_combinaciones = []


def x_clientes_cercanos(cliente, matriz_minutos, criterio = "Normal", clase = "", lista_clientes = clientes_faltantes, lista_clases = clases_clientes_faltantes):
    
    lista_clientes = [x for x in lista_clientes if x not in clientes_gold]
    clientes = filtro_clase_2(clase, lista_clientes)
    clientes = filtro_importe(criterio, clientes)
    listado = copy.deepcopy(matriz_minutos[lista_de_cliente.index(cliente)])
    nuevo_listado = [listado[lista_de_cliente.index(a)] for a in clientes]
    nuevo_listado.sort()
    nuevo_listado = nuevo_listado[0:6] #me falta parametrizarlo
    clientes_cercanos = [cliente]+[lista_de_cliente[listado.index(a)] for a in nuevo_listado]
    
    return clientes_cercanos

def filtro_clase_2(clase, lista_clientes):
    global lista_de_cliente
    global lista_de_clases
    
    nueva_lista_clientes = []
    lista_clases = [lista_de_clases[lista_de_cliente.index(x)] for x in lista_clientes]
    if clase == "":
        nueva_lista_clientes = copy.deepcopy(lista_clientes)
    else:
        for i,a in enumerate(lista_clientes):
            if diccionario_de_clases[clase] >= diccionario_de_clases[lista_clases[i]]:
                nueva_lista_clientes.append(a)
    return nueva_lista_clientes

def filtro_importe(criterio, lista_clientes):
    #global importe_alto
    nueva_lista_clientes = []
    if criterio == "Normal" or "":
        nueva_lista_clientes = copy.deepcopy(lista_clientes)
    elif criterio == "Importe bajo":
        for i,a in enumerate(lista_clientes):
            if listado_importes[lista_de_cliente.index(a)] < importe_bajo:
                nueva_lista_clientes.append(a)
    elif criterio == "Importe alto":
        for i,a in enumerate(lista_clientes):
            if listado_importes[lista_de_cliente.index(a)] > importe_alto:
                nueva_lista_clientes.append(a)
        
    return nueva_lista_clientes

#cambiar valores de los filtros (y ponerlos parametrizados)
def funcion_filtro_importes(ruta):
    importe = [x[5] for x in ruta]
    max_importe = max(importe)
    min_importe = min(importe)
    if min_importe < filtro_importe_ruta:
        min_importe = filtro_importe_ruta
    if max_importe < 3400:
        resultado = [x for x in ruta if x[5] >= max_importe-400]
    else:
        resultado = [x for x in ruta if x[5] >= 3000 and x[5] <= min_importe]
        if len(resultado) == 0:
            resultado = [x for x in ruta if x[5] >= 3000 and x[5] <= 4500]
        
    return resultado

def funcion_tiempo_al_almacen(listado_direcciones,direccion_almacen):
    tiempo_al_almacen = []
    for a in direcciones:
        tiempo_al_almacen.append(distancia(direccion_almacen,a)[0])
    return tiempo_al_almacen

def funcion_tiempo_al_almacen_random (listado_direcciones,direccion_almacen):
    tiempo_al_almacen = []
    for a in direcciones:
        tiempo_al_almacen.append(random.random() * 100)
    return tiempo_al_almacen

def funcion_tiempo_inicial(combinacion, tiempoalmacen):
    return tiempoalmacen[lista_de_cliente.index(combinacion[0])]

def funcion_tiempo_desplazamiento(combinacion,matriz_dir):
    tiempo = 0
    if len(combinacion) != 1:
        for i,a in enumerate(combinacion):
            if i != 0:
                tiempo += matriz_dir[lista_de_cliente.index(combinacion[i-1])][lista_de_cliente.index(a)][0]
    return tiempo

def funcion_tiempo_desplazamiento_km(combinacion,matriz_dir):
    km = 0
    if len(combinacion) != 1:
        for i,a in enumerate(combinacion):
            if i != 0:
                km += matriz_dir[lista_de_cliente.index(combinacion[i-1])][lista_de_cliente.index(a)][1]
    return km

def funcion_tiempo_dividido(combinacion, tiempoalmacen, matriz_dir, tiempo_al_centro):
    global lista_de_clases
    global lista_de_cliente
    
    tiempo_montaje = 0
    importe = 0
    if len(combinacion) == 1:
        tiempo_montaje += lista_de_tiempos_de_montaje[lista_de_cliente.index(combinacion[0])]
        importe += listado_importes[lista_de_cliente.index(combinacion[0])]
    else:
        for i,a in enumerate(combinacion):
            tiempo_montaje += lista_de_tiempos_de_montaje[lista_de_cliente.index(a)]
            importe += listado_importes[lista_de_cliente.index(a)]
            
    ultimo_cliente = combinacion[len(combinacion)-1]
    ultimo_cliente_centro = tiempo_al_centro[lista_de_cliente.index(ultimo_cliente)]
    if ultimo_cliente_centro > parametro_cliente_final:
        tiempo_desplazamiento = funcion_tiempo_desplazamiento(combinacion,matriz_dir) + ultimo_cliente_centro - parametro_cliente_final
    else:
        tiempo_desplazamiento = funcion_tiempo_desplazamiento(combinacion,matriz_dir)
    km_ruta = funcion_tiempo_desplazamiento_km(combinacion,matriz_dir)
    tiempo_al_almacen=funcion_tiempo_inicial(combinacion, tiempoalmacen)
    clasess = [lista_de_clases[lista_de_cliente.index(x)] for x in combinacion]
    categorias_montajes = [lista_categoria_montaje[lista_de_cliente.index(x)] for x in combinacion]
    n_montajes = 0
    exojo = 0
    cocina = 0
    rincon = 0
    for x in categorias_montajes:
        if x != 'T':
            n_montajes += 1
        if x == 'E':
            exojo += 1
        elif x == 'C':
            cocina += 1
        elif x == 'R':
            rincon += 1
    
    if 'S' in clasess:
        clase = 'S'
    elif 'A' in clasess:
        clase = 'A'
    elif 'B' in clasess:
        clase = 'B'
    elif 'N' in clasess:
        clase = 'N'
    elif 'T' in clasess:
        clase = 'T'
    elif 'X' in clasess:
        clase = 'T'
        
    gold = "NO"
    for a in combinacion:
        if a in clientes_gold:
            gold = "SI"
    
    return round(tiempo_montaje),combinacion,round(tiempo_desplazamiento),round(tiempo_al_almacen), len(combinacion), round(importe), clase, gold, n_montajes, exojo, cocina, rincon, round(km_ruta)

def funcion_suma_combinaciones(lista_resultado, lista_base, tiempoalmacen, matriz_dir, tiempo_al_centro):    
    global total_combinaciones
    calcula_tiempo = funcion_tiempo_dividido(lista_resultado, tiempoalmacen,matriz_dir, tiempo_al_centro)
    if calcula_tiempo not in total_combinaciones:
        total_combinaciones.append(calcula_tiempo)
    
    for a in lista_base:
        combinaciones = []
        combinaciones += lista_resultado
        dentro = 0
        if a not in lista_resultado:
            combinaciones.append(a)
            calcula_tiempo = funcion_tiempo_dividido(combinaciones, tiempoalmacen,matriz_dir, tiempo_al_centro)
            if calcula_tiempo[0]+calcula_tiempo[2]+calcula_tiempo[3] <= max_minutos_montaje and calcula_tiempo[8] <= n_maximo_montajes:
                total_combinaciones.append(calcula_tiempo)
                combinaciones = funcion_suma_combinaciones(combinaciones,lista_base, tiempoalmacen,matriz_dir, tiempo_al_centro)
            
    return total_combinaciones

def funcion_combinaciones_de_clientes_cercanos(clientes_cercanos_1, clienteee, tiempoalmacen, matriz_dir, tiempo_al_centro):
    global total_combinaciones
    total_combinaciones = []
    for a in clientes_cercanos_1:
        primero = []
        primero.append(a)
        resultado = funcion_suma_combinaciones(primero,clientes_cercanos_1, tiempoalmacen, matriz_dir, tiempo_al_centro)
    
    resultado = [x for x in resultado if clienteee in x[1]]
    return resultado

def distancia (origen, destino):
    ruta = gmaps.directions(origen,destino,mode="driving")
    return (ruta[0]['legs'][0]['duration']['value']/60,ruta[0]['legs'][0]['distance']['value']/1000)

def funcion_matriz_direcciones (direcciones):
    
    matriz_direcciones = []
    for i,origen in enumerate(direcciones):
        clientess = []
        for j,destino in enumerate(direcciones):
            if (i == j):
                clientess.append((1000.0,1000.0))
            else:
                clientess.append(distancia(origen,destino))
        matriz_direcciones.append(clientess)
    return matriz_direcciones

def funcion_matriz_direcciones_random (direcciones):
    
    matriz_direcciones = []
    for i,origen in enumerate(direcciones):
        clientess = []
        for j,destino in enumerate(direcciones):
            if (i == j):
                clientess.append((1000.0,1000.0))
            else:
                clientess.append((random.random() * 100, random.random() * 100))
        matriz_direcciones.append(clientess)
    return matriz_direcciones

def funcion_matriz_minutos(matriz_direcciones):
    matriz_minutos = []
    matriz_km = []
    for a in matriz_direcciones:
        fila_minutos = []
        fila_km = []
        for b in a:
            fila_minutos.append(b[0])
            fila_km.append(b[1])
        matriz_minutos.append(fila_minutos)
        matriz_km.append(fila_km)
    
    return matriz_minutos,matriz_km

def quita_duplicado(ruta):
    ruta_resultado = []
    combi_clientes = []
    for a in ruta:
        if set(a[1]) not in combi_clientes:
            combi_clientes.append(set(a[1]))
            ruta_resultado.append(a)
        else:
            indice = combi_clientes.index(set(a[1]))
            if ruta_resultado[indice][2]+ruta_resultado[indice][3] > a[2]+a[3]:
                ruta_resultado[indice] = a
    return ruta_resultado

def filtra_hora_rutas_gold(cliente,hora_gold,rutas,turno, tiempo_al_almacen, matriz_direcciones, tiempo_al_centro):
    
    rutas_filtradas = []
    
    if turno == 1:
        tiempo_al_turno = (hora_gold - 7.5)*60
    else:
        tiempo_al_turno = (hora_gold - 8.5)*60
    for a in rutas:
        indice2 = lista_de_cliente.index(cliente)
        indice = a[1].index(cliente)
        ruta_hasta_gold = a[1][0:indice+1]
        
        opcion = funcion_tiempo_dividido(ruta_hasta_gold,tiempo_al_almacen, matriz_direcciones, tiempo_al_centro)
        tiempo = opcion[0]+opcion[2]+opcion[3]-lista_de_tiempos_de_montaje[indice2]
        
        if tiempo <= tiempo_al_turno:
            rutas_filtradas.append(a)
    
    return rutas_filtradas

def funcion_calcula_x_ruta(n_rutas,tiempo_al_centro,matriz_minutos,tiempo_al_almacen,matriz_direcciones, n_clientes_listos, n_clientes_faltantes, clase):
    
    n_rutas_listas = []
    
    n_clientes_faltantes = filtro_clase_2(clase, n_clientes_faltantes)
    for i in range(n_rutas):
        tiempo_al_centro_2 = [tiempo_al_centro[j] for j,x in enumerate(lista_de_cliente) if x in n_clientes_faltantes and diccionario_de_clases[lista_de_clases[j]] <= diccionario_de_clases[clase]]
        tiempo_al_centro_2.sort()
        impor_siguientespivotes = tiempo_al_centro_2[len(tiempo_al_centro_2)-1]
        client_siguientespivotes = lista_de_cliente[tiempo_al_centro.index(impor_siguientespivotes)]
        clientes_cercanos = x_clientes_cercanos(client_siguientespivotes,matriz_minutos,"Normal",clase,n_clientes_faltantes)
        resultadooooooo = funcion_combinaciones_de_clientes_cercanos( clientes_cercanos, client_siguientespivotes, tiempo_al_almacen, matriz_direcciones, tiempo_al_centro)
        resultadooooooo = funcion_filtro_importes(resultadooooooo)
        cliente_1 = quita_duplicado(resultadooooooo)
        
        min_montaje = [x[0] for x in cliente_1]
        mejorrrr = max(min_montaje)
        seleccionada = cliente_1[min_montaje.index(mejorrrr)]
        
        while seleccionada[0]+seleccionada[2]+seleccionada[3] < 480:
            
            n_clientes = len(seleccionada[1])
            a = seleccionada[1][n_clientes-1]
            n_faltantes = [x for x in n_clientes_faltantes if x not in seleccionada[1]]
            
            if seleccionada[5] > 2000:
                listado = x_clientes_cercanos(a, matriz_minutos, "Importe bajo", clase, n_faltantes)
            else:
                listado = x_clientes_cercanos(a, matriz_minutos, "Importe alto", clase, n_faltantes)
            
            listado.pop(0)
            
            seleccionada = anadiendo_cliente(seleccionada, listado, tiempo_al_almacen, matriz_direcciones, "Normal", clase, tiempo_al_centro)
            if len(seleccionada[1]) == n_clientes:
                break
            else:
                if len(seleccionada[1]) <= 8 and seleccionada[7] == 'NO':
                    resultado = []
                    for a in list(permutations(seleccionada[1])):
                        resultado.append(funcion_tiempo_dividido(a,tiempo_al_almacen,matriz_direcciones, tiempo_al_centro))
                    tiempos = [x[0]+x[2]+x[3] for x in resultado]
                    seleccionada = resultado[tiempos.index(min(tiempos))]
        
        repe = [x for x in seleccionada[1] if x not in n_clientes_faltantes]
    
        if len(repe) == 0:
            nueva_ruta = funcion_tiempo_dividido(seleccionada[1], tiempo_al_almacen, matriz_direcciones, tiempo_al_centro)
            n_rutas_listas.append(nueva_ruta)

            n_clientes_listos += nueva_ruta[1]
            n_clientes_faltantes = [x for x in n_clientes_faltantes if x not in nueva_ruta[1]]
        
        #n_clientes_listos += seleccionada[1]
        #n_clientes_faltantes = [x for x in lista_de_cliente if x not in n_clientes_listos]
        
        #n_rutas_listas.append(seleccionada[0])
        
    return n_clientes_listos, n_clientes_faltantes, n_rutas_listas

def anadiendo_cliente(ruta, listaaaaa, tiempoalm, matr_dir, criterio, clase, tiempo_al_centro):
    global max_minutos_montaje
    
    listaaaaa = filtro_clase_2(clase, listaaaaa)
    listaaaaa = filtro_importe(criterio, listaaaaa)
    resultado = ruta
    for a in listaaaaa:
        clieeentes = copy.deepcopy(ruta[1])
        clieeentes.append(a)
        nueva_ruta = funcion_tiempo_dividido(clieeentes, tiempoalm, matr_dir, tiempo_al_centro)
        if nueva_ruta[0]+nueva_ruta[2]+nueva_ruta[3] <= max_minutos_montaje and nueva_ruta[8] <= n_maximo_montajes:
            resultado = nueva_ruta
            break
    return resultado

def calcula_coordenadas():
    global direcciones
    
    latitud = []
    longitud = []
    for x in direcciones:
        geocode_result = gmaps.geocode(x)
        latitud.append(geocode_result[0]['geometry']['location'] ['lat'])
        longitud.append(geocode_result[0]['geometry']['location']['lng'])
        
    return latitud,longitud

def matriz_info_general(matriz_equipos):
    
    clases = ['S','A','B','N','T']
    categorias = ['R','C','E','M','T']
    
    matriz_clases = []
    matriz_categorias = []
    
    matriz_clases_equipos = []
    matriz_categorias_equipos = []
    
    for a in clases:
        linea = []
        linea.append(a)
        linea.append(len([q for q in lista_de_clases if q == a]))
        linea.append(round(sum([lista_de_tiempos_de_montaje[i] for i,q in enumerate(lista_de_clases) if q == a])))
        linea.append(round(sum([listado_importes[i] for i,q in enumerate(lista_de_clases) if q == a])))
        
        linea2 = []
        linea2.append(a)
        linea2.append(len([x[9] for x in matriz_equipos if x[9] == a and x[1] == 'OK']))
        
        matriz_clases.append(linea)
        matriz_clases_equipos.append(linea2)
    
    linea = []
    linea.append('TOTAL')
    linea.append(len(lista_de_clases))
    linea.append(round(sum([q for q in lista_de_tiempos_de_montaje])))
    linea.append(round(sum([q for q in listado_importes])))
    
    linea2 = []
    linea2.append('TOTAL')
    linea2.append(len([x[9] for x in matriz_equipos if x[1] == 'OK']))
    
    matriz_clases.append(linea)
    matriz_clases_equipos.append(linea2)
    
    for a in categorias:
        linea = []
        linea.append(a)
        linea.append(len([q for q in lista_categoria_montaje if q == a]))
        linea.append(round(sum([lista_de_tiempos_de_montaje[i] for i,q in enumerate(lista_categoria_montaje) if q == a])))
        linea.append(round(sum([listado_importes[i] for i,q in enumerate(lista_categoria_montaje) if q == a])))
        
        linea2 = []
        linea2.append(a)
        #linea2.append(len([x[9] for x in matriz_equipos if x[9] == a]))
        
        matriz_categorias.append(linea)
        matriz_categorias_equipos.append(linea2)
    
    linea = []
    linea.append('TOTAL')
    linea.append(len(lista_de_clases))
    linea.append(round(sum([q for q in lista_de_tiempos_de_montaje])))
    linea.append(round(sum([q for q in listado_importes])))
    
    linea2 = []
    linea2.append('TOTAL')
    linea2.append(len([x[9] for x in matriz_equipos if x[1] == 'OK']))
    
    matriz_categorias.append(linea)
    matriz_categorias_equipos.append(linea2)
    
    return matriz_clases,matriz_categorias,matriz_clases_equipos,matriz_categorias_equipos

def leer_listado_equipos(archivo):
    wb = xlrd.open_workbook(archivo)
    
    hoja = wb.sheet_by_index(0)

    matriz = []
    i = 1
    while hoja.cell_value(i, 0) != '':
        fila = []
        for j in range(17):
            fila.append(hoja.cell_value(i, j))
        matriz.append(fila)
        i += 1
    
    return matriz

def funcion_matriz_para_avances(clientes_lejos,clientes_listos,rutas_listas,rutas_guardadas,equipos_disponibles,matriz_equipos,matriz_parametros_por_clases):
    matriz = []
    
    matriiiiizzz = []
    for j in range(4):
        categ_r = [x[4+j] for x in matriz_parametros_por_clases]
        for i,a in enumerate(categ_r):
            if a == 'SI':
                categ_r[i] = 1
            else:
                categ_r[i] = 0
        matriiiiizzz.append(categ_r)
        if j == 2:
            matriiiiizzz.append([1,1,1,1,0,1])
            matriiiiizzz.append([1,1,1,1,1,1])
        
    #totales
    clientes_ya_listos = len(clientes_listos)
    clientes_totales = len(lista_de_cliente)
    porcentaje_clientes = 0
    if clientes_totales != 0:
        porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
    rutas_ya_listas = len(rutas_listas) + len(rutas_guardadas)
    rutas_totales = len(equipos_disponibles)
    porcentaje_rutas = 0
    if rutas_totales != 0:
        porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
    linea = ['Clientes',clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
    matriz.append(linea)
    
    clientes_ya_listos = sum([lista_de_tiempos_de_montaje[lista_de_cliente.index(x)] for x in clientes_listos])
    clientes_totales = sum(lista_de_tiempos_de_montaje)
    porcentaje_clientes = 0
    if clientes_totales != 0:
        porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
    rutas_ya_listas = 0
    rutas_totales = 0
    porcentaje_rutas = 0
    if rutas_totales != 0:
        porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
    linea = ['Minutos',clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
    matriz.append(linea)
    
    clientes_ya_listos = round(sum([listado_importes[lista_de_cliente.index(x)] for x in clientes_listos]))
    clientes_totales = round(sum(listado_importes))
    porcentaje_clientes = 0
    if clientes_totales != 0:
        porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
    rutas_ya_listas = 0
    rutas_totales = 0
    porcentaje_rutas = 0
    if rutas_totales != 0:
        porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
    linea = ['Importe',clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
    matriz.append(linea)
    
    clllases = ['S','A','B','N','T']
    for x in clllases:
        clientes_ya_listos = len([a for a in clientes_listos if  lista_de_clases[lista_de_cliente.index(a)] == x])
        clientes_totales = len([a for i,a in enumerate(lista_de_cliente) if  lista_de_clases[i] == x])
        porcentaje_clientes = 0
        if clientes_totales != 0:
            porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
        rutas_ya_listas = rutas_listas + rutas_guardadas
        rutas_ya_listas = len([a for a in rutas_ya_listas if a[6] == x])
        
        rutas_totales = len([a[9] for a in matriz_equipos if a[9] == x and a[1] == 'OK'])
        porcentaje_rutas = 0
        if rutas_totales != 0:
            porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
        linea = [x,clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
        matriz.append(linea)
        
    clllases = ['R','C','E','M','T']
    for i,x in enumerate(clllases):
        clientes_ya_listos = len([a for a in clientes_listos if  lista_categoria_montaje[lista_de_cliente.index(a)] == x])
        clientes_totales = len([a for i,a in enumerate(lista_de_cliente) if  lista_categoria_montaje[i] == x])
        porcentaje_clientes = 0
        if clientes_totales != 0:
            porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
        rutas_ya_listas = rutas_listas + rutas_guardadas
        if i < 4:
            rutas_ya_listas = len([a for a in rutas_ya_listas if a[11-i] > 0])
        else:
            rutas_ya_listas = len(rutas_ya_listas)
        rutas_totales = matriiiiizzz[i][0] * matriz[3][5] + matriiiiizzz[i][1] * matriz[4][5] + matriiiiizzz[i][2] * matriz[5][5] + matriiiiizzz[i][3] * matriz[6][5] + matriiiiizzz[i][4] * matriz[7][5]
        
        porcentaje_rutas = 0
        if rutas_totales != 0:
            porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
        linea = [x,clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
        matriz.append(linea)
    
    clientes_ya_listos = 0
    clientes_totales = 0
    porcentaje_clientes = 0
    if clientes_totales != 0:
        porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
    rutas_ya_listas = 0
    rutas_totales = matriiiiizzz[5][0] * matriz[3][5] + matriiiiizzz[5][1] * matriz[4][5] + matriiiiizzz[5][2] * matriz[5][5] + matriiiiizzz[5][3] * matriz[6][5] + matriiiiizzz[5][4] * matriz[7][5]
    porcentaje_rutas = 0
    if rutas_totales != 0:
        porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
    linea = ['Pladur',clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
    matriz.append(linea)
    
    clientes_ya_listos = 0
    clientes_totales = 0
    porcentaje_clientes = 0
    if clientes_totales != 0:
        porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
    rutas_ya_listas = 0
    rutas_totales = 0
    porcentaje_rutas = 0
    if rutas_totales != 0:
        porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
    linea = ['Furgoneta peq',clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
    matriz.append(linea)
    
    clientes_ya_listos = 0
    clientes_totales = 0
    porcentaje_clientes = 0
    if clientes_totales != 0:
        porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
    rutas_ya_listas = 0
    rutas_totales = 0
    porcentaje_rutas = 0
    if rutas_totales != 0:
        porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
    linea = ['Z.Central',clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
    matriz.append(linea)
    
    clientes_ya_listos = len([x for x in clientes_listos if x in clientes_lejos])
    clientes_totales = len(clientes_lejos)
    porcentaje_clientes = 0
    if clientes_totales != 0:
        porcentaje_clientes = round(clientes_ya_listos/clientes_totales*100,1)
    rutas_ya_listas = 0
    rutas_totales = 0
    porcentaje_rutas = 0
    if rutas_totales != 0:
        porcentaje_rutas = round(rutas_ya_listas/rutas_totales*100,1)
    linea = ['F.Zona',clientes_ya_listos,clientes_totales,porcentaje_clientes,rutas_ya_listas,rutas_totales,porcentaje_rutas]
    matriz.append(linea)
    
    #--------------------------------------------------------------------------------------------------------------
    #---------------------------MATRIZ_2---------------------------------------------------------------------------
    matriz_2 = []
    matriz_2.append(['','Objetivo','Rutas Preparadas','Faltantes'])
    
    rutas_ya_listas = rutas_listas + rutas_guardadas
    
    importe_a_pagar = 0
    clientes_equipo = 0
    calculo_km = 0
    calculo_total = 0
    calculo_conducion = 0
    calculo_porcentaje = 0
    total_montaje = 0
    total_importe_bruto = 0
    total_importe_neto = 0
    for a in rutas_ya_listas:
        total_montaje += a[0]
        total_importe_bruto += a[5]
        total_importe_neto += a[5]
        aaa = round(a[5]*0.1)
        if aaa < 250:
            importe_a_pagar += 250
        else:
            importe_a_pagar += aaa
        clientes_equipo += a[4]
        calculo_km += a[12]
        calculo_total += a[0]+a[2]+a[3]
        calculo_conducion += a[2]+a[3]
    
    if calculo_total > 0:
        calculo_porcentaje = round(calculo_conducion * 100 / calculo_total,1)
    
    calculo_porcentaje_importe = 0
    if total_importe_neto > 0:
        calculo_porcentaje_importe = round(importe_a_pagar * 100 / total_importe_neto,1)
    
    if len(rutas_ya_listas) > 0:
        clientes_equipo = clientes_equipo/len(rutas_ya_listas)
        calculo_km = calculo_km/len(rutas_ya_listas)
    
    clientes_faltantes_equipo = 0
    clientes_medio_objetivo = len(lista_de_cliente)/len(equipos_disponibles)
    importe_bruto_ruta_objetivo = sum(listado_importes)/len(equipos_disponibles)
    
    total_importe_bruto_faltante = 0
    if len(rutas_ya_listas) < len(equipos_disponibles):
        clientes_faltantes_equipo = (len(lista_de_cliente)-len(clientes_listos))/(len(equipos_disponibles)-len(rutas_ya_listas))
        total_importe_bruto_faltante = (sum(listado_importes)-total_importe_bruto)/(len(equipos_disponibles)-len(rutas_ya_listas))
        total_montaje_faltante = (sum(lista_de_tiempos_de_montaje)-total_montaje)/(len(equipos_disponibles)-len(rutas_ya_listas))
    
    total_montaje_objetivo = sum(lista_de_tiempos_de_montaje)/len(equipos_disponibles)
    if len(rutas_ya_listas) > 0:
        total_montaje = total_montaje/len(rutas_ya_listas)
        total_importe_bruto = total_importe_bruto/len(rutas_ya_listas)
    
    
    min_a_pagar = len(equipos_disponibles) * 250
    if sum(listado_importes)*0.1 > min_a_pagar:
        min_a_pagar = sum(listado_importes)*0.1
    
    porcentaje_rutas_objetivo = round(min_a_pagar/sum(listado_importes)*100,1)
    
    matriz_2.append(['Minutos montaje',round(total_montaje_objetivo),round(total_montaje),round(total_montaje_faltante)])
    matriz_2.append(['Clientes medios',clientes_medio_objetivo,round(clientes_equipo,1),round(clientes_faltantes_equipo,1)])
    matriz_2.append(['Importe € neto',round(importe_bruto_ruta_objetivo),round(total_importe_bruto),round(total_importe_bruto_faltante)])
    matriz_2.append(['Coste €',round(min_a_pagar),round(importe_a_pagar),""])
    matriz_2.append(['Coste %',porcentaje_rutas_objetivo,calculo_porcentaje_importe,""])
    matriz_2.append(['% Drive time',18,calculo_porcentaje,""])
    matriz_2.append(['km',80,round(calculo_km),""])
    
    return matriz,matriz_2

def datos_cliente_x(cliente):
    global matriz_cliente_sin_duplicados
    global matriz_cliente
    
    cod_cliente = [a[0] for a in matriz_cliente_sin_duplicados]
    codigo_cliente = cod_cliente[lista_de_cliente.index(cliente)]
    
    lineas = [a for a in matriz_cliente if a[0] == codigo_cliente]
    
    return lineas