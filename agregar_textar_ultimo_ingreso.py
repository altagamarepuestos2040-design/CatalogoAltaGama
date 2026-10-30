# -*- coding: utf-8 -*-
"""
Agrega 6 articulos TEXTAR del ultimo ingreso (2026-10-30).

brakebook.com (la fuente de scraper_textar_brakebook.py) quedo detras de una
verificacion de Cloudflare, asi que los datos salen de las fichas publicas de
spareto.com (datos TecDoc): numeros OE, medidas, modelos y foto del fabricante.
El lado de 92319005/92319105 sale del numero OE de Porsche (301 izq / 302 der).

Uso:
    python agregar_textar_ultimo_ingreso.py
"""
import os, sys, json, requests

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARTICULOS_JSON = os.path.join(CARPETA, 'articulos.json')
IMG_DIR = os.path.join(CARPETA, 'Imagenes', 'Textar')
CDN = 'https://cdn.spareto.com/variants/images/'

PASTILLAS_MB = 'Juego de pastillas de freno'
DISCO = 'Disco de freno'


def modelos(marca, lista):
    return {marca: [{'m': m, 's': '', 'sub': ''} for m in lista]}


NUEVOS = [
    dict(id='2425301', desc=PASTILLAS_MB, lado='', img='001/297/180/original/converted-20230219-418386-tt8k32.jpg',
         oe={'MERCEDES-BENZ': '0004200205 | 0004209404 | 0054200720 | 0054204220 | 0054204320 | 005420432041 | 0054208120 | 0064201320 | 0064206120 | 0074206020 | 007420602090 | 0074208520'},
         compat=modelos('Mercedes-Benz', ['CLASE C, 2007 -', 'CLASE E, 2009 - 2016']),
         extra={'alto': '54,3', 'ancho': '122,5', 'espesor': '17,5', 'peso': '1,03', 'avisador': 'preparado para indicador de desgaste'}),
    dict(id='2637401', desc=PASTILLAS_MB, lado='', img='002/744/384/original/converted-20230219-418386-1ijbx92.jpg',
         oe={'MERCEDES-BENZ': '0004204103 | 0004204203 | 0004209305 | 0004209405 | 0004209505'},
         compat=modelos('Mercedes-Benz', ['CLASE C, 2021 -', 'CLASE E, 2023 -', 'GLC, 2022 -', 'CLASE S, 2020 -']),
         extra={'alto': '62', 'ancho': '117,1', 'espesor': '16,4', 'peso': '1,6', 'avisador': 'preparado para indicador de desgaste'}),
    dict(id='2644301', desc=PASTILLAS_MB, lado='', img='002/744/396/original/converted-20230219-418386-bx16k9.jpg',
         oe={'MERCEDES-BENZ': '0004200604 | 0004200704 | 0004203603 | 0004203703 | 0004213102 | 0004215402'},
         compat=modelos('Mercedes-Benz', ['CLASE E, 2023 -', 'GLC, 2022 -', 'CLASE S, 2020 -']),
         extra={'alto': '98,7', 'ancho': '138,6', 'espesor': '17,9', 'peso': '2,84', 'avisador': 'preparado para indicador de desgaste'}),
    dict(id='92275605', desc=DISCO, lado='', img='002/745/125/original/converted-20230219-418386-xzb3gl.jpg',
         oe={'MINI': '34119804829 | 34119811538 | 9804829 | 9811538'},
         compat=modelos('MINI', ['COUNTRYMAN (R60), 2010 - 2016', 'PACEMAN (R61), 2012 - 2016']),
         extra={'alto': '45,3', 'espesor': '24', 'peso': '7,8', 'min_qty': 2, 'mult': 2}),
    dict(id='92319005', desc=DISCO, lado='Eje delantero', img='002/745/300/original/converted-20230219-418386-u380z6.jpg',
         oe={'PORSCHE': '95B615301 | 95B615301AB | 95B615301F | 95B615301H | 95B615301T'},
         compat=modelos('Porsche', ['MACAN (95B), 2014 -']),
         extra={'alto': '52,2', 'espesor': '34', 'peso': '12,2', 'min_qty': 1, 'colocacion': 'Delantero izquierdo',
                'juego': '92319105', 'juego_principal': True}),
    dict(id='92319105', desc=DISCO, lado='Eje delantero', img='002/745/319/original/converted-20230219-418386-1dwb9ii.jpg',
         oe={'PORSCHE': '95B615302 | 95B615302AB | 95B615302F | 95B615302H | 95B615302R'},
         compat=modelos('Porsche', ['MACAN (95B), 2014 -']),
         extra={'alto': '52,2', 'espesor': '34', 'peso': '12,2', 'min_qty': 1, 'colocacion': 'Delantero derecho',
                'juego': '92319005', 'juego_secundario': True}),
]


def main():
    arts = json.load(open(ARTICULOS_JSON, encoding='utf-8'))
    existentes = {a['id'] for a in arts}
    for n in NUEVOS:
        if n['id'] in existentes:
            print(n['id'], 'ya estaba')
            continue
        r = requests.get(CDN + n['img'], timeout=30, headers={'User-Agent': 'Mozilla/5.0'})
        img = ''
        if r.ok and len(r.content) > 1500:
            img = f"Imagenes/Textar/{n['id']}.jpg"
            open(os.path.join(CARPETA, img), 'wb').write(r.content)
        art = {
            'id': n['id'], 'marca': 'TEXTAR', 'desc': n['desc'], 'cat': 'Frenos', 'tipo': '',
            'lado': n['lado'], 'props': {}, 'fabs': ', '.join(n['oe']), 'oe': n['oe'], 'rel': '',
            'url': '', 'img': img, 'compat': n['compat'],
        }
        art.update(n['extra'])
        arts.append(art)
        print(f"{n['id']:>9} | {n['desc']:<28} | {img or 'SIN FOTO'}")
    with open(ARTICULOS_JSON, 'w', encoding='utf-8') as f:
        json.dump(arts, f, ensure_ascii=False, separators=(',', ':'))
    print('Total:', len(arts))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
