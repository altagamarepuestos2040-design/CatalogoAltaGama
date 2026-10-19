# -*- coding: utf-8 -*-
"""
Agrega al catalogo Alta Gama los articulos febi del Excel "nuevo ingreso Febi"
que todavia no estan en articulos.json.

Datos: API de partsfinder.bilsteingroup.com (necesita la cookie JOURNEY_ID que
el sitio da al abrir cualquier pagina; el script la consigue solo).
Fotos: CDN pf-article-details (las de buena calidad), {id}_img1.jpg .. _img6.jpg.

El sitio corta con 429/403 si se consulta rapido: va de a uno con pausa y guarda
cada respuesta en cache_febi/, asi se puede cortar y retomar.

Uso:
    python agregar_febi_nuevo_ingreso.py            (baja datos y fotos)
    python agregar_febi_nuevo_ingreso.py --agregar  (ademas agrega a articulos.json)
"""
import os, sys, json, time, requests, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

CARPETA = os.path.dirname(os.path.abspath(__file__))
EXCEL = r'C:\Users\Usuario\Documents\nuevo ingreso Febi.xlsx'
ARTICULOS_JSON = os.path.join(CARPETA, 'articulos.json')
IMG_DIR = os.path.join(CARPETA, 'Imagenes', 'Febi')
CACHE = os.path.join(CARPETA, 'cache_febi')
PAUSA = 4

API = ('https://partsfinder.bilsteingroup.com/api/articles/{}?include=images'
       '&filter%5BvehicleType%5D=CAR&filter%5Bcountry%5D=AR'
       '&filter%5Bbrands%5D=FEBI%2CSWAG%2CBLUE_PRINT%2CMOTAIR')
CDN = 'https://cdn.partsfinder.bilsteingroup.com/pf-article-details/1/{}_{}.jpg'
HEADERS = {
    'Accept': 'application/vnd.api+json',
    'Accept-Language': 'es-ES,es;q=0.9',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36',
    'Referer': 'https://partsfinder.bilsteingroup.com/es/',
}

# Mismas categorias que ya usa el catalogo (de scraper_febi_nuevos.py)
CATEGORY_MAP = [
    ('Refrigeración', ['depósito compensador', 'refriger', 'radiador', 'termostato', 'ventilador', 'calefacción', 'enfriamiento', 'enfriador', 'bomba de agua', 'intercambiador']),
    ('Frenos', ['freno', 'pastilla', 'zapata', 'forros', 'abrasión', 'desgaste']),
    ('Embrague', ['embrague', 'desembrague', 'plato de presión']),
    ('Distribución / Correas', ['correa', 'tensor', 'polea', 'kit de distribución', 'cadena de distribución', 'guía', 'carril', 'rueda dentada', 'piñón']),
    ('Transmisión', ['cardan', 'cardán', 'árbol de transmisión', 'junta homocinética', 'caja de cambios', 'diferencial', 'transmisión', 'semieje', 'engranaje', 'fuelle', 'palanca de cambios']),
    ('Suspensión / Dirección', ['rodamiento', 'rótula', 'buje', 'silentblock', 'brazo', 'amortiguador', 'muelle', 'resorte', 'mangueta', 'estabilizador', 'barra de acoplamiento', 'tirante', 'articulación', 'dirección', 'acoplamiento', 'soporte de apoyo', 'tope', 'apoyo', 'suspensión', 'terminal', 'cubo de rueda', 'perno']),
    ('Sensores / Electrónica', ['sensor', 'sonda', 'interruptor', 'caudalímetro', 'presostato', 'actuador', 'bobina', 'electroválvula', 'relé', 'fusible', 'transmisor', 'cable']),
    ('Lubricantes y Fluidos', ['aceite', 'lubricante', 'líquido', 'fluido', 'grasa', 'aditivo', 'anticongelante']),
    ('Juntas', ['junta', 'retén', 'sello', 'anillo', 'obturador']),
    ('Bombas', ['bomba']),
    ('Mantenimiento', ['filtro', 'bujía', 'tapón', 'tapa de cierre', 'tapa', 'escobilla', 'kit de servicio']),
    ('Motor', ['soporte', 'motor', 'cigüeñal', 'árbol de levas', 'válvula', 'culata', 'pistón', 'cilindro', 'colector', 'manguera', 'tubo', 'ventilación']),
]


def categorizar(desc):
    d = desc.lower()
    for cat, claves in CATEGORY_MAP:
        if any(k in d for k in claves):
            return cat
    return 'Motor'


def ids_faltantes():
    with open(ARTICULOS_JSON, encoding='utf-8') as f:
        arts = json.load(f)
    en_cat = {str(a['id']).lstrip('0') for a in arts if str(a.get('marca', '')).lower() == 'febi'}
    ws = openpyxl.load_workbook(EXCEL, data_only=True)['Hoja1']
    vistos, res = set(), []
    for (v,) in ws.iter_rows(max_col=1, values_only=True):
        if v in (None, ''):
            continue
        n = str(v).strip().lstrip('0')
        if n and n not in en_cat and n not in vistos:
            vistos.add(n)
            res.append(n.zfill(5))
    return arts, res


def bajar_datos(ids):
    os.makedirs(CACHE, exist_ok=True)
    s = requests.Session()
    s.headers.update(HEADERS)

    def nueva_sesion():
        # La API solo responde con la cookie que da el sitio al abrir una pagina
        s.cookies.clear()
        s.get('https://partsfinder.bilsteingroup.com/es/article/febi/173613',
              headers={'Accept': 'text/html'}, timeout=20)

    nueva_sesion()
    for i, art_id in enumerate(ids, 1):
        destino = os.path.join(CACHE, art_id + '.json')
        if os.path.exists(destino):
            continue
        espera = 60
        while True:
            r = s.get(API.format(art_id), timeout=20)
            if r.status_code == 401:
                nueva_sesion()
                r = s.get(API.format(art_id), timeout=20)
            if r.status_code in (403, 429):
                print(f'  bloqueado ({r.status_code}), espero {espera}s...', flush=True)
                time.sleep(espera)
                espera = min(espera * 2, 900)
                continue
            break
        if r.status_code == 404:
            json.dump({'no_existe': True}, open(destino, 'w', encoding='utf-8'))
            print(f'[{i}/{len(ids)}] {art_id}: no existe en febi')
        elif r.ok:
            open(destino, 'w', encoding='utf-8').write(r.text)
            print(f'[{i}/{len(ids)}] {art_id}: ok', flush=True)
        else:
            print(f'[{i}/{len(ids)}] {art_id}: error {r.status_code}')
        time.sleep(PAUSA)


def bajar_fotos(art_id):
    """Baja _1.._6 del CDN de detalle; devuelve cuantas hay."""
    n = 0
    for k in range(1, 7):
        destino = os.path.join(IMG_DIR, f'{art_id}_img{k}.jpg')
        if os.path.exists(destino):
            n = k
            continue
        r = requests.get(CDN.format(art_id, k), timeout=20,
                         headers={'User-Agent': HEADERS['User-Agent'], 'Referer': 'https://partsfinder.bilsteingroup.com/'})
        if r.status_code != 200 or len(r.content) < 1500:
            break
        open(destino, 'wb').write(r.content)
        n = k
    return n


def armar_articulo(art_id, j):
    a = j['data']['attributes']
    oe = {}
    for e in a.get('oeNumbers', []):
        if e.get('make') and e.get('numbers'):
            oe[e['make']] = ', '.join(e['numbers'])
    props = {}
    for at in a.get('articleAttributes', []):
        val = str(at.get('value', '')) + (' ' + at['unit'] if at.get('unit') else '')
        props[at['type']] = props[at['type']] + ' / ' + val if at['type'] in props else val
    desc = (a.get('articleDescription') or '').strip()
    return {
        'id': art_id,
        'marca': 'Febi',
        'desc': desc,
        'cat': categorizar(desc),
        'tipo': 'Kit completo' if (a.get('packagingQty') or 1) > 1 else 'Pieza suelta',
        'lado': (a.get('fittingSide') or '').strip(),
        'props': props,
        'fabs': ', '.join(oe.keys()),
        'oe': oe,
        'rel': '',
        'url': f'https://partsfinder.bilsteingroup.com/es/article/febi/{art_id}',
        'img': f'Imagenes/Febi/{art_id}_img1.jpg',
        'compat': {},
    }


def main():
    arts, ids = ids_faltantes()
    print(f'Faltan en el catalogo: {len(ids)}')
    bajar_datos(ids)
    if '--agregar' not in sys.argv:
        return
    nuevos, sin_datos, sin_foto, no_existe = [], [], [], []
    for art_id in ids:
        p = os.path.join(CACHE, art_id + '.json')
        if not os.path.exists(p):
            sin_datos.append(art_id)
            continue
        j = json.load(open(p, encoding='utf-8'))
        if j.get('no_existe'):
            no_existe.append(art_id)
            continue
        art = armar_articulo(art_id, j)
        if bajar_fotos(art_id) == 0:
            sin_foto.append(art_id)
        nuevos.append(art)
    arts.extend(nuevos)
    with open(ARTICULOS_JSON, 'w', encoding='utf-8') as f:
        json.dump(arts, f, ensure_ascii=False, separators=(',', ':'))
    print(f'Agregados: {len(nuevos)} | sin datos: {sin_datos} | no existen: {no_existe} | sin foto: {sin_foto}')


if __name__ == '__main__':
    main()
