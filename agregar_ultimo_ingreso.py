# -*- coding: utf-8 -*-
"""
Agrega al catalogo los articulos del Excel "precios de ultimo ingreso" que no
estaban en articulos.json (2026-10-29). Los precios EUR ya se cargaron aparte
en precios_eur.json.

- Con datos en partsfinder: se arman igual que en agregar_febi_nuevo_ingreso.py.
- Codigos febi discontinuados (el sitio redirige a un reemplazo): se agregan con
  su propio codigo y numero original del Excel, datos y fotos del reemplazo, y el
  reemplazo en 'rel'.
- Sin datos en ningun lado: descripcion y numero original del Excel, sin foto.

Uso:
    python agregar_ultimo_ingreso.py
"""
import os, sys, json, shutil
import agregar_febi_nuevo_ingreso as febi

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARTICULOS_JSON = os.path.join(CARPETA, 'articulos.json')
IMG_DIR = os.path.join(CARPETA, 'Imagenes', 'Febi')

# codigo -> (marca, numero original del Excel, descripcion del Excel)
CON_DATOS = ['185446', 'ADBP210027']
REEMPLAZADOS = {  # codigo viejo -> reemplazo (ya esta en el catalogo)
    '31468': ('30696', '8K0407515'),
    '26082': ('36735', '1K0407365C'),
    '26083': ('36736', '1K0407366C'),
    '33153': ('33152', '059145201F'),
    '104204': ('22932', '1K0199555S'),
}
SIN_DATOS = [
    ('25332', 'Febi', '059115124C', 'Tensor de cadena de distribución'),
    ('105774', 'Febi', '5010359706', 'Filtro de combustible'),
    ('02-37475-02', 'Victor Reinz', '06H198012', 'Juego de juntas de culata'),
]


def cache(art_id):
    return json.load(open(os.path.join(febi.CACHE, art_id + '.json'), encoding='utf-8'))


def main():
    arts = json.load(open(ARTICULOS_JSON, encoding='utf-8'))
    existentes = {a['id'] for a in arts}
    nuevos = []

    for art_id in CON_DATOS:
        art = febi.armar_articulo(art_id, cache(art_id))
        if art_id.startswith('AD'):
            art['marca'] = 'Blue Print'
            art['url'] = f'https://partsfinder.bilsteingroup.com/es/article/blueprint/{art_id}'
            art['cat'] = 'Mantenimiento'  # carcasa de filtro: "aceite" la mandaba a Lubricantes
        if febi.bajar_fotos(art_id) == 0:
            art['img'] = ''
        nuevos.append(art)

    for viejo, (reemplazo, oe_excel) in REEMPLAZADOS.items():
        art = febi.armar_articulo(reemplazo, cache(reemplazo))
        art['id'] = viejo
        art['url'] = f'https://partsfinder.bilsteingroup.com/es/article/febi/{viejo}'
        art['rel'] = reemplazo
        art['oe'] = {'Nº original': oe_excel, **art['oe']}
        n = febi.bajar_fotos(reemplazo)
        for k in range(1, n + 1):
            shutil.copy2(os.path.join(IMG_DIR, f'{reemplazo}_img{k}.jpg'),
                         os.path.join(IMG_DIR, f'{viejo}_img{k}.jpg'))
        art['img'] = f'Imagenes/Febi/{viejo}_img1.jpg' if n else ''
        nuevos.append(art)

    for art_id, marca, oe_excel, desc in SIN_DATOS:
        nuevos.append({
            'id': art_id, 'marca': marca, 'desc': desc,
            'cat': 'Juntas' if marca == 'Victor Reinz' else febi.categorizar(desc),
            'tipo': 'Kit completo' if desc.lower().startswith('juego') else 'Pieza suelta',
            'lado': '', 'props': {}, 'fabs': '', 'oe': {'Nº original': oe_excel},
            'rel': '', 'url': '', 'img': '', 'compat': {},
        })

    agregados = [a for a in nuevos if a['id'] not in existentes]
    arts.extend(agregados)
    with open(ARTICULOS_JSON, 'w', encoding='utf-8') as f:
        json.dump(arts, f, ensure_ascii=False, separators=(',', ':'))
    for a in agregados:
        print(f"{a['id']:>12} | {a['marca']:<12} | {a['cat']:<24} | {a['img'] or 'sin foto'} | {a['desc']}")
    print(f'Agregados: {len(agregados)} | total: {len(arts)}')


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
