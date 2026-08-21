# -*- coding: utf-8 -*-
"""
Agrega los articulos de Victor Reinz (descargados en Desktop/Victor Reinz) al
catalogo Alta Gama (articulos.json de este repo), con compatibilidad de
vehiculo completa (modelo/anio/motorizacion) para que el buscador por auto
del sitio los encuentre.

Fuente: Desktop/Victor Reinz/progreso.json (SKU -> imagenes ya descargadas)
Vuelve a consultar la API de TecAlliance por cada SKU para traer OE completo
y el detalle de aplicacion de vehiculo (no solo el resumen por marca).

Sin precios (se agregan despues aparte).

Uso:
    python agregar_victor_reinz.py

Genera/actualiza:
    articulos.json          (agrega los articulos nuevos, no pisa los existentes)
    Imagenes/Victor Reinz/  (copia las imagenes ya descargadas)
    progreso_altagama.json  (que SKU ya se agregaron, para poder cortar y retomar)
"""
import os
import re
import json
import time
import shutil
import requests

CARPETA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
CARPETA_VR_DESKTOP = r'C:\Users\Usuario\Desktop\Victor Reinz'
IMAGENES_VR_ORIGEN = os.path.join(CARPETA_VR_DESKTOP, 'imagenes')
PROGRESO_VR = os.path.join(CARPETA_VR_DESKTOP, 'progreso.json')

ARTICULOS_JSON = os.path.join(CARPETA_SCRIPT, 'articulos.json')
IMAGENES_DESTINO = os.path.join(CARPETA_SCRIPT, 'Imagenes', 'Victor Reinz')
PROGRESO_ALTAGAMA = os.path.join(CARPETA_SCRIPT, 'progreso_victor_reinz_altagama.json')

BASE_WS = 'https://webservice.tecalliance.services'
ENDPOINT = f'{BASE_WS}/pegasus-3-0/services/TecdocToCatDLB.jsonEndpoint'
PROVIDER = 2044
DATA_SUPPLIER_ID = 9  # Victor Reinz dentro del catalogo TecAlliance
FILTER_QUERIES = ['(dataSupplierId NOT IN (4978,4982))']
PAUSA = 0.25

# Marca TecAlliance (manuName) -> valor exacto del <select id="fveh"> del sitio.
# Los que no tienen equivalente relevante para venta de autos/utilitarios en
# Argentina (camiones pesados, buses, tuners) se dejan afuera a proposito.
MARCA_TECALLIANCE_A_SITIO = {
    'ALFA ROMEO': 'Alfa Romeo',
    'AUDI': 'Audi',
    'AUDI (FAW)': 'Audi-FAW',
    'BENTLEY': 'Bentley',
    'BMW': 'BMW',
    'BMW (BRILLIANCE)': 'BMW-Brilliance',
    'CHERY': 'Chery',
    'CHEVROLET': 'Chevrolet',
    'CHRYSLER': 'Chrysler',
    'CITROËN': 'Citroën',
    'CITROEN': 'Citroën',
    'CUPRA': 'CUPRA',
    'DACIA': 'Dacia',
    'DAEWOO': 'Daewoo',
    'DODGE': 'Dodge',
    'FIAT': 'FIAT',
    'FORD': 'Ford',
    'GEELY': 'Geely (Jili)',
    'GREAT WALL': 'Great Wall',
    'HAVAL': 'Haval',
    'HONDA': 'Honda',
    'HYUNDAI': 'Hyundai',
    'JAGUAR': 'Jaguar',
    'JEEP': 'Jeep',
    'KIA': 'KIA',
    'KTM': 'KTM',
    'LADA': 'Lada',
    'LAMBORGHINI': 'Lamborghini',
    'LANCIA': 'Lancia',
    'LAND ROVER': 'Land Rover',
    'LEXUS': 'Lexus',
    'MASERATI': 'Maserati',
    'MAZDA': 'Mazda',
    'MERCEDES-BENZ': 'Mercedes-Benz',
    'MERCEDES-BENZ (BBDC)': 'Beijing-Benz-Chrysler (BBDC)',
    'MERCEDES-BENZ (FJDA)': 'Fujian Daimler',
    'MINI': 'MINI',
    'MITSUBISHI': 'Mitsubishi',
    'NISSAN': 'Nissan',
    'OPEL': 'Opel',
    'PEUGEOT': 'Peugeot',
    'PORSCHE': 'Porsche',
    'PROTON': 'Proton',
    'PUCH': 'Puch (Steyr)',
    'RENAULT': 'Renault',
    'ROVER': 'Rover',
    'SAAB': 'Saab',
    'SEAT': 'Seat',
    'SKODA': 'Škoda',
    'SMART': 'Smart',
    'SSANGYONG': 'Ssangyong',
    'SUBARU': 'Subaru',
    'SUZUKI': 'Suzuki',
    'TOYOTA': 'Toyota',
    'VOLVO': 'Volvo',
    'VW': 'Volkswagen (VW)',
    'VW (FAW)': 'Volkswagen (VW)-FAW',
}

CAT_POR_DEFECTO = 'Juntas'
TIPO_POR_DEFECTO = 'Pieza suelta'


def crear_sesion():
    s = requests.Session()
    s.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Content-Type': 'application/json',
        'Accept': 'application/json',
    })
    return s


def llamar_api(s, metodo, params, reintentos=3):
    for intento in range(reintentos):
        try:
            r = s.post(ENDPOINT, data=json.dumps({metodo: params}), timeout=30)
            if r.status_code != 200:
                time.sleep(1)
                continue
            data = r.json()
            if isinstance(data, dict) and data.get('status') not in (None, 200):
                return None
            return data
        except (requests.RequestException, ValueError):
            time.sleep(1)
    return None


def buscar_articulo(s, sku):
    body = {
        'applyDqmRules': True,
        'articleCountry': 'DE',
        'provider': PROVIDER,
        'lang': 'es',
        'searchQuery': sku,
        'searchMatchType': 'exact',
        'searchType': 0,
        'page': 1,
        'perPage': 10,
        'filterQueries': FILTER_QUERIES,
        'dataSupplierIds': [DATA_SUPPLIER_ID],
        'includeGenericArticles': True,
        'includeImages': True,
        'includeOEMNumbers': True,
        'includeArticleText': True,
    }
    data = llamar_api(s, 'getArticles', body)
    if not data:
        return None
    for a in data.get('articles', []) or []:
        if a.get('articleNumber', '').upper() == sku.upper():
            return a
    articulos = data.get('articles', []) or []
    return articulos[0] if articulos else None


def obtener_marcas_aplicacion(s, article_id, tipo):
    data = llamar_api(s, 'getArticleLinkedAllLinkingTargetManufacturer2', {
        'provider': PROVIDER, 'articleId': article_id, 'articleCountry': 'DE',
        'country': 'ALL', 'countryGroupFlag': True, 'lang': 'es', 'linkingTargetType': tipo,
    })
    if not data:
        return []
    return (data.get('data') or {}).get('array', []) or []


def obtener_pares_marca(s, article_id, manu_id, tipo):
    data = llamar_api(s, 'getArticleLinkedAllLinkingTarget4', {
        'provider': PROVIDER, 'articleId': article_id, 'lang': 'es', 'articleCountry': 'DE',
        'country': 'ALL', 'countryGroupFlag': True, 'linkingTargetManuId': manu_id,
        'linkingTargetType': tipo, 'withMainArticles': False,
    })
    if not data:
        return []
    pares = []
    for grupo in (data.get('data') or {}).get('array', []) or []:
        linkages = grupo.get('articleLinkages')
        if not isinstance(linkages, dict):
            continue
        for link in linkages.get('array', []) or []:
            pares.append({'articleLinkId': link['articleLinkId'], 'linkingTargetId': link['linkingTargetId']})
    return pares


def formatear_anio(yyyymm):
    if not yyyymm:
        return ''
    s = str(yyyymm)
    return s[:4] if len(s) >= 4 else s


def formatear_potencia(desde, hasta):
    if not desde:
        return ''
    if hasta and hasta != desde:
        return f'{desde}-{hasta}'
    return str(desde)


def resolver_vehiculos(s, article_id, pares, tamano_lote=25):
    """Devuelve lista de dicts {m, s} unicos a partir de pares articleLink/linkingTarget (tipo VOL)."""
    vistos = set()
    filas = []
    for i in range(0, len(pares), tamano_lote):
        lote = pares[i:i + tamano_lote]
        data = llamar_api(s, 'getArticleLinkedAllLinkingTargetsByIds3', {
            'provider': PROVIDER, 'lang': 'es', 'articleCountry': 'DE', 'articleId': article_id,
            'immediateAttributs': True, 'linkingTargetType': 'VOL',
            'linkedArticlePairs': {'array': lote},
        })
        if not data:
            continue
        for item in (data.get('data') or {}).get('array', []) or []:
            vehiculos = item.get('linkedVehicles')
            if not isinstance(vehiculos, dict):
                continue
            for v in vehiculos.get('array', []) or []:
                modelo = v.get('modelDesc', '')
                anio_desde = formatear_anio(v.get('yearOfConstructionFrom'))
                anio_hasta = formatear_anio(v.get('yearOfConstructionTo'))
                m = f'{modelo} {anio_desde} - {anio_hasta}'.strip() if anio_hasta else f'{modelo} {anio_desde} -'.strip()

                car = v.get('carDesc', '')
                kw = formatear_potencia(v.get('powerKwFrom'), v.get('powerKwTo'))
                hp = formatear_potencia(v.get('powerHpFrom'), v.get('powerHpTo'))
                partes_s = [car]
                if kw or hp:
                    partes_s.append(f'{kw} kW / {hp} hp'.strip())
                s_txt = ' - '.join(p for p in partes_s if p)

                clave = (m, s_txt)
                if clave in vistos:
                    continue
                vistos.add(clave)
                filas.append({'m': m, 's': s_txt})
        time.sleep(0.15)
    return filas


def obtener_compat_completo(s, article_id):
    """Devuelve (compat_dict, fabs_texto) usando solo marcas mapeadas al sitio."""
    marcas_vol = obtener_marcas_aplicacion(s, article_id, 'VOL')
    compat = {}
    for marca in marcas_vol:
        nombre_site = MARCA_TECALLIANCE_A_SITIO.get(marca['manuName'])
        if not nombre_site:
            continue  # marca de camion/bus/tuner no vendida como auto en el sitio
        pares = obtener_pares_marca(s, article_id, marca['manuId'], 'VOL')
        if not pares:
            continue
        filas = resolver_vehiculos(s, article_id, pares)
        if filas:
            compat.setdefault(nombre_site, [])
            compat[nombre_site].extend(filas)
        time.sleep(0.15)
    fabs = ', '.join(compat.keys())
    return compat, fabs


def extraer_oe(art):
    oe = {}
    for o in art.get('oemNumbers', []) or []:
        marca_raw = o.get('mfrName', '')
        numero = o.get('articleNumber', '')
        if not numero:
            continue
        nombre_site = MARCA_TECALLIANCE_A_SITIO.get(marca_raw, marca_raw)
        oe.setdefault(nombre_site, [])
        if numero not in oe[nombre_site]:
            oe[nombre_site].append(numero)
    return {k: ', '.join(v) for k, v in oe.items()}


def cargar_json(ruta, default):
    if os.path.exists(ruta):
        with open(ruta, 'r', encoding='utf-8') as f:
            return json.load(f)
    return default


def guardar_json(ruta, data):
    with open(ruta, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))


def main():
    progreso_vr = cargar_json(PROGRESO_VR, {})
    skus = [sku for sku, archivos in progreso_vr.items() if archivos]
    print(f'{len(skus)} SKUs de Victor Reinz con imagen para agregar')

    articulos = cargar_json(ARTICULOS_JSON, [])
    ids_existentes = set(str(a['id']) for a in articulos)
    print(f'articulos.json actual: {len(articulos)} articulos')

    os.makedirs(IMAGENES_DESTINO, exist_ok=True)
    progreso_ag = cargar_json(PROGRESO_ALTAGAMA, {})

    s = crear_sesion()
    agregados = 0
    ya_estaban = 0
    sin_datos = []

    for idx, sku in enumerate(skus, 1):
        if sku in ids_existentes:
            ya_estaban += 1
            continue
        if progreso_ag.get(sku) == 'ok':
            continue

        print(f'[{idx}/{len(skus)}] {sku}...', end=' ', flush=True)
        art = buscar_articulo(s, sku)
        if not art:
            print('SIN DATOS')
            sin_datos.append(sku)
            time.sleep(PAUSA)
            continue

        genericos = art.get('genericArticles', []) or []
        descripcion = genericos[0].get('genericArticleDescription', '') if genericos else ''
        article_id = genericos[0].get('legacyArticleId') if genericos else None
        if len(genericos) > 1:
            # El articulo tiene mas de una clasificacion generica (ej: selladores /
            # productos multiuso que TecDoc lista bajo varias categorias de junta).
            # La descripcion tomada (la primera) puede no ser la correcta -> revisar a mano.
            otras = [g.get('genericArticleDescription', '') for g in genericos[1:]]
            print(f'  [AVISO] {sku}: clasificacion ambigua, revisar descripcion manualmente. '
                  f'Elegida: "{descripcion}" | Otras: {otras}')

        oe = extraer_oe(art)
        compat, fabs = ({}, '') if not article_id else obtener_compat_completo(s, article_id)

        # Copiar imagen ya descargada
        archivos_origen = progreso_vr.get(sku, [])
        nombre_img_destino = ''
        for fname in archivos_origen:
            src = os.path.join(IMAGENES_VR_ORIGEN, fname)
            dst = os.path.join(IMAGENES_DESTINO, fname)
            if os.path.exists(src):
                shutil.copy2(src, dst)
                nombre_img_destino = fname
                break

        nuevo = {
            'id': sku,
            'marca': 'Victor Reinz',
            'desc': descripcion,
            'cat': CAT_POR_DEFECTO,
            'tipo': TIPO_POR_DEFECTO,
            'lado': '',
            'fabs': fabs,
            'oe': oe,
            'rel': '',
            'url': f'https://web.tecalliance.net/reinz/es/parts/{DATA_SUPPLIER_ID}/{sku}/detail?query={sku}',
            'img': f'Imagenes/Victor Reinz/{nombre_img_destino}' if nombre_img_destino else '',
            'compat': compat,
        }
        articulos.append(nuevo)
        ids_existentes.add(sku)
        agregados += 1
        progreso_ag[sku] = 'ok'

        total_vehs = sum(len(v) for v in compat.values())
        print(f'OK -> {len(compat)} marcas, {total_vehs} aplicaciones - {descripcion}')

        if idx % 10 == 0:
            guardar_json(ARTICULOS_JSON, articulos)
            guardar_json(PROGRESO_ALTAGAMA, progreso_ag)
            print(f'  [Guardado parcial {idx}/{len(skus)}]')

        time.sleep(PAUSA)

    guardar_json(ARTICULOS_JSON, articulos)
    guardar_json(PROGRESO_ALTAGAMA, progreso_ag)

    print()
    print('=' * 60)
    print(f'Agregados: {agregados}')
    print(f'Ya estaban en el catalogo: {ya_estaban}')
    print(f'Sin datos: {len(sin_datos)} {sin_datos}')
    print(f'Total articulos.json: {len(articulos)}')


if __name__ == '__main__':
    main()
