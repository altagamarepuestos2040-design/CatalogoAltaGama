# -*- coding: utf-8 -*-
"""
Agrega articulos BOGAP nuevos al catalogo Alta Gama, usando la API oficial
de bogap.de (la misma que ya usan descargar_imagenes_bogap.py y
actualizar_desc_bogap.py: POST bogap.de/BOGAPAPP/ProductHandler.aspx,
action=GetInfo, bogapNo={codigo}).

Fuente de codigos: Desktop/Bogap/bogap.xlsx (columna "NUMERO BOGAP"), filtrando
los que ya estan en articulos.json.

Sin precios (se agregan despues aparte).

Uso:
    python agregar_bogap_nuevos.py

Genera/actualiza:
    articulos.json          (agrega los articulos nuevos, no pisa los existentes)
    Imagenes/Bogap/         (descarga fotos: {id}.jpg + {id}_img2.jpg.._img6.jpg)
    progreso_bogap_nuevos.json  (que codigos ya se procesaron, para poder cortar y retomar)
"""
import os
import json
import time
import requests
import openpyxl

CARPETA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
EXCEL_BOGAP = r'C:\Users\Usuario\Desktop\Bogap\bogap.xlsx'

ARTICULOS_JSON = os.path.join(CARPETA_SCRIPT, 'articulos.json')
IMAGENES_DESTINO = os.path.join(CARPETA_SCRIPT, 'Imagenes', 'Bogap')
PROGRESO = os.path.join(CARPETA_SCRIPT, 'progreso_bogap_nuevos.json')

API_URL = 'https://www.bogap.de/BOGAPAPP/ProductHandler.aspx'
PAUSA = 0.5

# Marca TecDoc (manuName, en applicationList) -> valor exacto del <select id="fveh">
# del sitio. Mismo criterio que agregar_victor_reinz.py: se dejan afuera las
# marcas de camion/bus/tuner que no se venden como auto en el sitio.
MARCA_A_SITIO = {
    'ALFA ROMEO': 'Alfa Romeo', 'AUDI': 'Audi', 'AUDI (FAW)': 'Audi-FAW',
    'BENTLEY': 'Bentley', 'BMW': 'BMW', 'BMW (BRILLIANCE)': 'BMW-Brilliance',
    'CHERY': 'Chery', 'CHEVROLET': 'Chevrolet', 'CHRYSLER': 'Chrysler',
    'CITROËN': 'Citroën', 'CITROEN': 'Citroën', 'CUPRA': 'CUPRA',
    'DACIA': 'Dacia', 'DAEWOO': 'Daewoo', 'DODGE': 'Dodge', 'FIAT': 'FIAT',
    'FORD': 'Ford', 'GEELY': 'Geely (Jili)', 'GREAT WALL': 'Great Wall',
    'HAVAL': 'Haval', 'HONDA': 'Honda', 'HYUNDAI': 'Hyundai', 'JAGUAR': 'Jaguar',
    'JEEP': 'Jeep', 'KIA': 'KIA', 'KTM': 'KTM', 'LADA': 'Lada',
    'LAMBORGHINI': 'Lamborghini', 'LANCIA': 'Lancia', 'LAND ROVER': 'Land Rover',
    'LEXUS': 'Lexus', 'MASERATI': 'Maserati', 'MAZDA': 'Mazda',
    'MERCEDES-BENZ': 'Mercedes-Benz', 'MERCEDES-BENZ (BBDC)': 'Beijing-Benz-Chrysler (BBDC)',
    'MERCEDES-BENZ (FJDA)': 'Fujian Daimler', 'MINI': 'MINI', 'MITSUBISHI': 'Mitsubishi',
    'NISSAN': 'Nissan', 'OPEL': 'Opel', 'PEUGEOT': 'Peugeot', 'PORSCHE': 'Porsche',
    'PROTON': 'Proton', 'PUCH': 'Puch (Steyr)', 'RENAULT': 'Renault', 'ROVER': 'Rover',
    'SAAB': 'Saab', 'SEAT': 'Seat', 'SKODA': 'Škoda', 'SMART': 'Smart',
    'SSANGYONG': 'Ssangyong', 'SUBARU': 'Subaru', 'SUZUKI': 'Suzuki', 'TOYOTA': 'Toyota',
    'VOLVO': 'Volvo', 'VW': 'Volkswagen (VW)', 'VOLKSWAGEN': 'Volkswagen (VW)',
    'VW (FAW)': 'Volkswagen (VW)-FAW',
}

# Traduccion ingles -> espanol de categoryName (heredado de actualizar_desc_bogap.py
# + las nuevas encontradas en la lista de codigos a agregar).
TRADUCCION = {
    'Gasket, cylinder head cover': 'Junta, tapa de la culata',
    'Gasket, cylinder head': 'Junta de culata',
    'Gasket Set, cylinder head': 'Juego de juntas, culata',
    'Gasket, intake manifold': 'Junta, colector de admisión',
    'Gasket, exhaust manifold': 'Junta, colector de escape',
    'Gasket, oil sump': 'Junta, cárter de aceite',
    'Gasket, thermostat housing': 'Junta, carcasa del termostato',
    'Gasket, water pump': 'Junta, bomba de agua',
    'Gasket, timing case cover': 'Junta, tapa de distribución',
    'Gasket, crankcase ventilation': 'Junta, ventilación del cárter',
    'Gasket, exhaust pipe': 'Junta, tubo de escape',
    'Gasket, oil cooler': 'Junta, enfriador de aceite',
    'Gasket, oil filter housing': 'Junta, carcasa del filtro de aceite',
    'Gasket Set': 'Juego de juntas',
    'Cylinder Head Cover': 'Tapa de la culata',
    'Cover, timing belt': 'Tapa de distribución',
    'Cover, timing chain': 'Tapa de cadena de distribución',
    'Oil Filler Cap': 'Tapón, llenado de aceite',
    'Oil Sump': 'Cárter de aceite',
    'Timing Belt': 'Correa de distribución',
    'Timing Belt Kit': 'Kit de distribución',
    'Timing Chain': 'Cadena de distribución',
    'Timing Chain Kit': 'Kit de cadena de distribución',
    'Guide, timing chain': 'Guía, cadena de distribución',
    'Tensioner, timing chain': 'Tensor, cadena de distribución',
    'Tensioner Lever, timing chain': 'Palanca tensora, cadena de distribución',
    'Tensioner, timing belt': 'Tensor, correa de distribución',
    'Deflection/Guide Pulley, timing belt': 'Polea tensora, correa de distribución',
    'Deflection Pulley/Guide Pulley, timing belt': 'Polea tensora, correa de distribución',
    'Deflection/Guide Pulley, V-ribbed belt': 'Polea tensora, correa poly-V',
    'Belt Tensioner, V-ribbed belt': 'Tensor, correa poly-V',
    'Belt Pulley, crankshaft': 'Polea, cigüeñal',
    'Belt Pulley, water pump': 'Polea, bomba de agua',
    'Sprocket, crankshaft': 'Piñón, cigüeñal',
    'Timing Case Cover': 'Tapa de la caja de distribución',
    'Chain Tensioner, oil pump drive': 'Tensor de cadena, accionamiento bomba de aceite',
    'Camshaft': 'Árbol de levas',
    'Camshaft Adjuster': 'Regulador del árbol de levas',
    'Control Valve, camshaft adjustment': 'Válvula de control, regulación árbol de levas',
    'Seal Cap, camshaft': 'Tapón de obturación, árbol de levas',
    'Seal, camshaft adjustment': 'Retén, regulación árbol de levas',
    'Bearing Bracket, camshaft': 'Soporte de cojinete, árbol de levas',
    'Tappet': 'Empujador de válvula',
    'Tappet, high-pressure pump': 'Empujador, bomba de alta presión',
    'Balance Shaft, crank drive': 'Eje de equilibrado, cigüeñal',
    'Connecting Rod Bearing Shell': 'Casquillo de biela',
    'Shaft Seal, crankshaft': 'Retén, cigüeñal',
    'Oil Sump, automatic transmission': 'Cárter, transmisión automática',
    'Oil Cooler, automatic transmission': 'Enfriador de aceite, transmisión automática',
    'Oil Cooler, engine oil': 'Enfriador de aceite de motor',
    'Oil Pressure Valve': 'Válvula de presión de aceite',
    'Oil Separator, crankcase ventilation': 'Separador de aceite, ventilación del cárter',
    'Baffle Plate, oil sump': 'Deflector, cárter de aceite',
    'Tube, oil dipstick': 'Varilla de nivel de aceite',
    'Seal, oil filter': 'Retén, filtro de aceite',
    'Seal, oil pump': 'Retén, bomba de aceite',
    'Sensor, engine oil level': 'Sensor, nivel de aceite del motor',
    'Cap, oil filler neck': 'Tapón de llenado de aceite',
    'Cap, oil filter housing': 'Tapón, carcasa del filtro de aceite',
    'Oil Line, charger': 'Tubería de aceite, turbo',
    'Water Pump, engine cooling': 'Bomba de agua, refrigeración del motor',
    'Thermostat, coolant': 'Termostato de refrigerante',
    'Cap, coolant tank': 'Tapón, depósito de refrigerante',
    'Coolant Flange': 'Brida de refrigerante',
    'Coolant Pipe, charger': 'Tubería de refrigerante, turbo',
    'Coolant Pipe': 'Tubería de refrigerante',
    'Fan, engine cooling': 'Ventilador, refrigeración del motor',
    'Control Unit, electric fan (engine cooling)': 'Módulo de control, ventilador eléctrico',
    'Mounting, radiator': 'Soporte del radiador',
    'Temperature Switch, radiator fan': 'Interruptor de temperatura, ventilador',
    'Auxiliary Water Pump (cooling water circuit)': 'Bomba de agua auxiliar (circuito de refrigeración)',
    'High-pressure Pump': 'Bomba de alta presión',
    'Relay, fuel pump': 'Relé, bomba de combustible',
    'Control Unit, fuel pump': 'Módulo de control, bomba de combustible',
    'Charcoal Filter, tank ventilation': 'Filtro de carbón activo, ventilación depósito',
    'Valve, charcoal filter (tank ventilation)': 'Válvula, filtro de carbón activo (ventilación depósito)',
    'Flange, fuel supply unit': 'Brida, unidad de suministro de combustible',
    'fuel supply unit': 'Unidad de suministro de combustible',
    'Actuator, turbocharger': 'Actuador, turbocompresor',
    'Boost Pressure Control Valve': 'Válvula de control de presión de sobrealimentación',
    'Intake Manifold Module': 'Módulo del colector de admisión',
    'Divert-air Valve, charger': 'Válvula de derivación de aire, turbo',
    'Seal Ring, charge air hose': 'Anillo obturador, manguera de aire de carga',
    'Sensor, boost pressure': 'Sensor, presión de sobrealimentación',
    'Sensor, suction pipe change-over flap': 'Sensor, trampilla del colector de admisión',
    'Valve, exhaust gas recirculation': 'Válvula EGR',
    'Sensor, exhaust gas temperature': 'Sensor, temperatura de gases de escape',
    'Oxygen Sensor': 'Sonda Lambda',
    'Sensor, wheel speed': 'Sensor, velocidad de rueda',
    'Sensor, intake air temperature': 'Sensor, temperatura del aire de admisión',
    'Sensor, fuel pressure': 'Sensor, presión de combustible',
    'Sensor, exterior temperature': 'Sensor, temperatura exterior',
    'Sensor, headlight levelling': 'Sensor, nivelación de faros',
    'Sensor, park distance control': 'Sensor, aparcamiento',
    'Mass Air Flow Sensor': 'Sensor de masa de aire (caudalímetro)',
    'Stop Light Switch': 'Interruptor de luz de freno',
    'Hazard Warning Light Switch': 'Interruptor de luces de emergencia',
    'Ignition Switch': 'Interruptor de encendido',
    'Switch, headlight': 'Interruptor de faros',
    'Switch, window regulator': 'Interruptor, elevalunas',
    'Switch, exterior mirror adjustment': 'Interruptor, ajuste espejo exterior',
    'Warning Contact, brake pad wear': 'Contacto de desgaste de pastillas',
    'Clock Spring, airbag': 'Resorte espiral, airbag (cinta volante)',
    'Control Unit, lights': 'Módulo de control, luces',
    'Bushing, stabiliser bar': 'Silent block, barra estabilizadora',
    'Bushing, axle beam': 'Silent block, viga del eje',
    'Mounting, control/trailing arm': 'Silent block, brazo de suspensión',
    'Dust Cover Kit, shock absorber': 'Kit fuelle/tope, amortiguador',
    'Suspension, propshaft': 'Soporte central, árbol de transmisión',
    'Joint Kit, drive shaft': 'Kit de junta homocinética',
    'Hydraulic Filter, automatic transmission': 'Filtro hidráulico, transmisión automática',
    'Hydraulic Filter, multi-plate clutch (all-wheel drive)': 'Filtro hidráulico, embrague multidisco (tracción integral)',
    'Hydraulic Pump, steering': 'Bomba de dirección hidráulica',
    'Interior Blower': 'Ventilador interior',
    'Heat Exchanger, interior heating': 'Intercambiador de calor, calefacción',
    'Control Element, heating/ventilation': 'Mando, climatización',
    'Actuator, blending flap': 'Actuador, compuerta de mezcla',
    'Compressor, compressed-air system': 'Compresor de aire comprimido',
    'Gas Spring, bonnet': 'Amortiguador de capó',
    'Gas Spring, boot/cargo area': 'Amortiguador de maletero',
    'Actuator, fuel filler flap': 'Actuador, tapa del depósito',
    'Bonnet Lock': 'Cerradura del capó',
    'Wiper Linkage': 'Mecanismo limpiaparabrisas',
    'Wiper Arm, window cleaning': 'Brazo limpiaparabrisas',
    'Wiper Arm Set, window cleaning': 'Juego de brazos limpiaparabrisas',
    'Wiper Blade': 'Escobilla limpiaparabrisas',
    'Washer Fluid Jet, window cleaning': 'Tobera lavaparabrisas',
    'Washer Fluid Jet, headlight cleaning': 'Tobera lavafaros',
    'Filter, cabin air': 'Filtro de habitáculo',
    'Repair Kit, crankcase ventilation': 'Kit de reparación, ventilación del cárter',
    'Repair Kit, water pump': 'Kit de reparación, bomba de agua',
    'Oil Filter': 'Filtro de aceite',
    'Oil Filter Housing': 'Carcasa del filtro de aceite',
    'Housing, oil filter': 'Carcasa, filtro de aceite',
    'Oil Pressure Switch': 'Presostato de aceite',
    'Oil Cooler': 'Enfriador de aceite',
    'Oil Pump': 'Bomba de aceite',
    'Oil Separator': 'Separador de aceite',
    'Valve, crankcase ventilation': 'Válvula, ventilación del cárter',
    'Hose, crankcase ventilation': 'Manguera, ventilación del cárter',
    'Water Pump': 'Bomba de agua',
    'Thermostat': 'Termostato',
    'Thermostat Housing': 'Carcasa del termostato',
    'Radiator, engine cooling': 'Radiador, refrigeración del motor',
    'Radiator Hose': 'Manguera del radiador',
    'Coolant Hose': 'Manguera de refrigerante',
    'Expansion Tank, coolant': 'Vaso de expansión, refrigerante',
    'Coolant Control Valve': 'Válvula de control de refrigerante',
    'Air Filter': 'Filtro de aire',
    'Intake Hose, air filter': 'Manguera de admisión, filtro de aire',
    'Intake Manifold': 'Colector de admisión',
    'Throttle Body': 'Cuerpo del acelerador',
    'Turbocharger': 'Turbocompresor',
    'Charge Air Hose': 'Manguera de aire de carga',
    'Fuel Filter': 'Filtro de combustible',
    'Fuel Pump': 'Bomba de combustible',
    'Fuel Rail': 'Rampa de inyectores',
    'Injector': 'Inyector',
    'Ignition Coil': 'Bobina de encendido',
    'Spark Plug': 'Bujía',
    'Sensor, crankshaft pulse': 'Sensor, señal del cigüeñal',
    'Sensor, camshaft position': 'Sensor, posición del árbol de levas',
    'Sensor, throttle position': 'Sensor, posición del acelerador',
    'Sensor, oil pressure': 'Sensor, presión de aceite',
    'Sensor, coolant temperature': 'Sensor, temperatura del refrigerante',
    'Lambda Sensor': 'Sonda Lambda',
    'Brake Disc': 'Disco de freno',
    'Brake Pad Set, disc brake': 'Juego de pastillas de freno',
    'Brake Drum': 'Tambor de freno',
    'Brake Shoe Set': 'Juego de zapatas de freno',
    'Brake Caliper': 'Pinza de freno',
    'Brake Master Cylinder': 'Cilindro maestro de freno',
    'Brake Hose': 'Flexible de freno',
    'Wheel Brake Cylinder': 'Cilindro de rueda',
    'Vacuum Pump, braking system': 'Bomba de vacío, sistema de frenos',
    'Clutch Kit': 'Kit de embrague',
    'Slave Cylinder, clutch': 'Bombín de embrague',
    'Master Cylinder, clutch': 'Cilindro maestro de embrague',
    'Clutch Cable': 'Cable de embrague',
    'Flywheel': 'Volante motor',
    'Hydraulic Filter Kit, automatic transmission': 'Kit de filtro hidráulico, transmisión automática',
    'Mounting, engine': 'Soporte de motor',
    'Mounting, automatic transmission': 'Soporte, transmisión automática',
    'Mounting, manual transmission': 'Soporte, caja de cambios manual',
    'Steering Pump': 'Bomba de dirección',
    'Power Steering Pump': 'Bomba de dirección asistida',
    'Tie Rod End': 'Extremo de rótula de dirección',
    'Tie Rod': 'Barra de dirección',
    'Steering Rack': 'Cremallera de dirección',
    'Link/Coupling Rod, stabiliser bar': 'Bieleta de barra estabilizadora',
    'Stabiliser Bar': 'Barra estabilizadora',
    'Control/Trailing Arm, wheel suspension': 'Brazo de suspensión',
    'Control Arm': 'Brazo de control',
    'Trailing Arm': 'Tirante de suspensión',
    'Ball Joint': 'Rótula de suspensión',
    'Wheel Bearing Kit': 'Kit de rodamiento de rueda',
    'Wheel Bearing': 'Rodamiento de rueda',
    'Wheel Hub': 'Buje de rueda',
    'Shock Absorber': 'Amortiguador',
    'Suspension Strut': 'Puntal de suspensión',
    'Suspension Strut Support Mount': 'Plato de amortiguador',
    'Air Spring, suspension': 'Resorte neumático, suspensión',
    'Protective Cap/Bellow, shock absorber': 'Fuelle/Tope, amortiguador',
    'Rubber Buffer, suspension': 'Tope de goma, suspensión',
    'Catalytic Converter': 'Catalizador',
    'Exhaust Pipe': 'Tubo de escape',
    'Silencer/Muffler': 'Silenciador',
    'EGR Valve': 'Válvula EGR',
    'Compressor, air conditioning': 'Compresor de aire acondicionado',
    'Condenser, air conditioning': 'Condensador de aire acondicionado',
    'Expansion Valve, air conditioning': 'Válvula de expansión, aire acondicionado',
    'Tailgate Lock': 'Cerradura del portón trasero',
    'Door Lock': 'Cerradura de puerta',
    'Window Regulator': 'Elevalunas',
    'Hydraulic Pump': 'Bomba hidráulica',
    'Hydraulic Filter': 'Filtro hidráulico',
    'Auxiliary Drive Belt': 'Correa auxiliar',
    'Tensioner, auxiliary drive belt': 'Tensor, correa auxiliar',
    'Deflection/Guide Pulley, auxiliary drive belt': 'Polea tensora, correa auxiliar',
}

# categoryName (ingles) -> categoria del sitio. Cubre las categorias que
# aparecen en TRADUCCION. Lo que no esta mapeado cae en 'General' y se
# avisa por consola para revisar a mano.
CAT_BUCKET = {}
def _bucket(cat, *categorias):
    for c in categorias:
        CAT_BUCKET[c] = cat

_bucket('Juntas',
    'Gasket, cylinder head cover', 'Gasket, cylinder head', 'Gasket Set, cylinder head',
    'Gasket, intake manifold', 'Gasket, exhaust manifold', 'Gasket, oil sump',
    'Gasket, thermostat housing', 'Gasket, water pump', 'Gasket, timing case cover',
    'Gasket, crankcase ventilation', 'Gasket, exhaust pipe', 'Gasket, oil cooler',
    'Gasket, oil filter housing', 'Gasket Set')

_bucket('Distribución / Correas',
    'Cover, timing belt', 'Cover, timing chain', 'Timing Belt', 'Timing Belt Kit',
    'Timing Chain', 'Timing Chain Kit', 'Guide, timing chain', 'Tensioner, timing chain',
    'Tensioner Lever, timing chain', 'Tensioner, timing belt',
    'Deflection/Guide Pulley, timing belt', 'Deflection Pulley/Guide Pulley, timing belt',
    'Deflection/Guide Pulley, V-ribbed belt', 'Belt Tensioner, V-ribbed belt',
    'Belt Pulley, crankshaft', 'Belt Pulley, water pump', 'Sprocket, crankshaft',
    'Timing Case Cover', 'Chain Tensioner, oil pump drive', 'Auxiliary Drive Belt',
    'Tensioner, auxiliary drive belt', 'Deflection/Guide Pulley, auxiliary drive belt')

_bucket('Motor',
    'Cylinder Head Cover', 'Oil Filler Cap', 'Oil Sump', 'Camshaft', 'Camshaft Adjuster',
    'Control Valve, camshaft adjustment', 'Seal Cap, camshaft', 'Seal, camshaft adjustment',
    'Bearing Bracket, camshaft', 'Tappet', 'Tappet, high-pressure pump',
    'Balance Shaft, crank drive', 'Connecting Rod Bearing Shell', 'Shaft Seal, crankshaft',
    'Oil Pressure Valve', 'Oil Separator, crankcase ventilation', 'Baffle Plate, oil sump',
    'Tube, oil dipstick', 'Seal, oil filter', 'Seal, oil pump',
    'Sensor, engine oil level', 'Cap, oil filler neck', 'Cap, oil filter housing',
    'Oil Line, charger', 'Boost Pressure Control Valve', 'Intake Manifold Module',
    'Divert-air Valve, charger', 'Seal Ring, charge air hose',
    'Valve, exhaust gas recirculation', 'Oxygen Sensor', 'Intake Manifold',
    'Throttle Body', 'Turbocharger', 'Charge Air Hose', 'Fuel Rail', 'Injector',
    'Valve, crankcase ventilation', 'Hose, crankcase ventilation',
    'Intake Hose, air filter', 'Actuator, turbocharger', 'EGR Valve', 'Catalytic Converter',
    'Exhaust Pipe', 'Silencer/Muffler', 'fuel supply unit',
    'Valve, charcoal filter (tank ventilation)', 'Charcoal Filter, tank ventilation',
    'Flange, fuel supply unit', 'Oil Filter', 'Oil Filter Housing', 'Housing, oil filter',
    'Air Filter', 'Fuel Filter', 'Repair Kit, crankcase ventilation')

_bucket('Refrigeración',
    'Water Pump, engine cooling', 'Thermostat, coolant', 'Cap, coolant tank',
    'Coolant Flange', 'Coolant Pipe, charger', 'Coolant Pipe', 'Fan, engine cooling',
    'Control Unit, electric fan (engine cooling)', 'Mounting, radiator',
    'Temperature Switch, radiator fan', 'Water Pump', 'Thermostat', 'Thermostat Housing',
    'Radiator, engine cooling', 'Radiator Hose', 'Coolant Hose',
    'Expansion Tank, coolant', 'Coolant Control Valve', 'Repair Kit, water pump')

_bucket('Bombas',
    'High-pressure Pump', 'Relay, fuel pump', 'Control Unit, fuel pump', 'Fuel Pump',
    'Oil Pump', 'Steering Pump', 'Power Steering Pump', 'Hydraulic Pump',
    'Hydraulic Pump, steering', 'Auxiliary Water Pump (cooling water circuit)',
    'Oil Cooler, automatic transmission', 'Oil Cooler, engine oil', 'Oil Cooler',
    'Vacuum Pump, braking system')

_bucket('Sensores / Electrónica',
    'Sensor, boost pressure', 'Sensor, suction pipe change-over flap', 'Sensor, wheel speed',
    'Sensor, intake air temperature', 'Sensor, fuel pressure', 'Sensor, exterior temperature',
    'Sensor, headlight levelling', 'Sensor, park distance control', 'Mass Air Flow Sensor',
    'Stop Light Switch', 'Hazard Warning Light Switch', 'Ignition Switch',
    'Switch, headlight', 'Switch, window regulator', 'Switch, exterior mirror adjustment',
    'Warning Contact, brake pad wear', 'Clock Spring, airbag', 'Control Unit, lights',
    'Ignition Coil', 'Spark Plug', 'Sensor, crankshaft pulse', 'Sensor, camshaft position',
    'Sensor, throttle position', 'Sensor, oil pressure', 'Sensor, coolant temperature',
    'Lambda Sensor', 'Sensor, exhaust gas temperature')

_bucket('Suspensión / Dirección',
    'Bushing, stabiliser bar', 'Bushing, axle beam', 'Mounting, control/trailing arm',
    'Dust Cover Kit, shock absorber', 'Tie Rod End', 'Tie Rod', 'Steering Rack',
    'Link/Coupling Rod, stabiliser bar', 'Stabiliser Bar', 'Control/Trailing Arm, wheel suspension',
    'Control Arm', 'Trailing Arm', 'Ball Joint', 'Wheel Bearing Kit', 'Wheel Bearing',
    'Wheel Hub', 'Shock Absorber', 'Suspension Strut', 'Suspension Strut Support Mount',
    'Air Spring, suspension', 'Protective Cap/Bellow, shock absorber', 'Rubber Buffer, suspension',
    'Mounting, engine')

_bucket('Transmisión',
    'Joint Kit, drive shaft', 'Hydraulic Filter, automatic transmission',
    'Hydraulic Filter, multi-plate clutch (all-wheel drive)', 'Hydraulic Filter Kit, automatic transmission',
    'Hydraulic Filter', 'Oil Sump, automatic transmission', 'Mounting, automatic transmission',
    'Mounting, manual transmission', 'Suspension, propshaft', 'Clutch Kit',
    'Slave Cylinder, clutch', 'Master Cylinder, clutch', 'Clutch Cable', 'Flywheel')

_bucket('Frenos',
    'Brake Disc', 'Brake Pad Set, disc brake', 'Brake Drum', 'Brake Shoe Set',
    'Brake Caliper', 'Brake Master Cylinder', 'Brake Hose', 'Wheel Brake Cylinder')

_bucket('Mantenimiento',
    'Filter, cabin air', 'Wiper Blade', 'Wiper Arm, window cleaning',
    'Wiper Arm Set, window cleaning', 'Washer Fluid Jet, window cleaning',
    'Washer Fluid Jet, headlight cleaning', 'Wiper Linkage')

_bucket('General',
    'Interior Blower', 'Heat Exchanger, interior heating', 'Control Element, heating/ventilation',
    'Actuator, blending flap', 'Compressor, compressed-air system', 'Gas Spring, bonnet',
    'Gas Spring, boot/cargo area', 'Actuator, fuel filler flap', 'Bonnet Lock',
    'Compressor, air conditioning', 'Condenser, air conditioning',
    'Expansion Valve, air conditioning', 'Tailgate Lock', 'Door Lock', 'Window Regulator')

# attrName (parammeterList) -> nombre de propiedad en espanol, tal cual usa
# el catalogo (ver ejemplo A1210228: "Longitud (mm)", "Ancho (mm)", "Altura", "Peso (kg)")
PROP_TRADUCCION = {
    'Length [mm]': 'Longitud (mm)',
    'Width [mm]': 'Ancho (mm)',
    'Height [mm]': 'Altura',
    'Weight [kg]': 'Peso (kg)',
    'Thickness [mm]': 'Espesor (mm)',
    'Diameter [mm]': 'Diámetro (mm)',
    'Inner Diameter [mm]': 'Diámetro interior (mm)',
    'Outer Diameter [mm]': 'Diámetro exterior (mm)',
}


def leer_codigos_nuevos():
    wb = openpyxl.load_workbook(EXCEL_BOGAP, data_only=True)
    ws = wb.worksheets[0]
    codigos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            c = str(row[0]).strip()
            if c and c not in codigos:
                codigos.append(c)

    with open(ARTICULOS_JSON, encoding='utf-8') as f:
        articulos = json.load(f)
    existentes = set(str(a['id']).strip().upper() for a in articulos if a.get('marca') == 'BOGAP')

    return [c for c in codigos if c.upper() not in existentes], articulos


def crear_sesion():
    s = requests.Session()
    s.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Referer': 'https://www.bogap.de/',
    })
    return s


def obtener_info(s, cod, reintentos=3):
    for intento in range(reintentos):
        try:
            r = s.post(API_URL, data={'action': 'GetInfo', 'bogapNo': cod}, timeout=20)
            if r.status_code != 200:
                time.sleep(1)
                continue
            d = r.json()
            if d.get('success') and d.get('data'):
                return d['data']
            return None
        except (requests.RequestException, ValueError):
            time.sleep(1)
    return None


def descargar_imagen(s, url, ruta):
    try:
        r = s.get(url, timeout=25)
        if r.status_code == 200 and len(r.content) > 2000:
            with open(ruta, 'wb') as f:
                f.write(r.content)
            return True
    except requests.RequestException:
        pass
    return False


def formatear_anio(yyyymm):
    if not yyyymm:
        return ''
    s = str(yyyymm)
    return s[:4] if len(s) >= 4 else s


def construir_compat(data):
    """Devuelve (compat_dict, fabs_texto) a partir de applicationList."""
    compat = {}
    for app in data.get('applicationList', []) or []:
        marca_raw = (app.get('manuName') or '').strip()
        nombre_site = MARCA_A_SITIO.get(marca_raw.upper())
        if not nombre_site:
            continue
        modelo = app.get('modelName', '')
        anio_desde = formatear_anio(app.get('yearOfConstrFrom'))
        anio_hasta = formatear_anio(app.get('yearOfConstrTo'))
        m = f'{modelo} {anio_desde} - {anio_hasta}'.strip() if anio_hasta else f'{modelo} {anio_desde} -'.strip()

        tipo = app.get('typeName', '')
        motor_codes = app.get('motorCodes', '')
        s_txt = tipo
        if motor_codes:
            s_txt = f'{tipo} ({motor_codes})' if tipo else motor_codes

        fila = {'m': m, 's': s_txt}
        compat.setdefault(nombre_site, [])
        if fila not in compat[nombre_site]:
            compat[nombre_site].append(fila)

    fabs = ', '.join(compat.keys())
    return compat, fabs


def construir_oe(data):
    numeros = []
    for x in data.get('replXRefList', []) or []:
        n = (x.get('oeNumber') or '').strip()
        if n and n not in numeros:
            numeros.append(n)
    if not numeros:
        return {}
    return {'OE': ', '.join(numeros)}


def construir_props(data):
    props = {}
    ancho = alto = peso = ''
    for p in data.get('parammeterList', []) or []:
        attr = p.get('attrName', '')
        val = p.get('attrValue', '')
        if not attr or not val:
            continue
        nombre_es = PROP_TRADUCCION.get(attr, attr)
        props[nombre_es] = val
        if attr == 'Width [mm]':
            ancho = val
        elif attr == 'Height [mm]':
            alto = val
        elif attr == 'Weight [kg]':
            peso = val
    return props, ancho, alto, peso


def main():
    codigos, articulos = leer_codigos_nuevos()
    print(f'{len(codigos)} codigos BOGAP nuevos para agregar')

    os.makedirs(IMAGENES_DESTINO, exist_ok=True)
    progreso = {}
    if os.path.exists(PROGRESO):
        with open(PROGRESO, encoding='utf-8') as f:
            progreso = json.load(f)

    s = crear_sesion()
    agregados = 0
    sin_datos = []
    sin_traducir = set()
    sin_categoria = set()

    for idx, cod in enumerate(codigos, 1):
        if progreso.get(cod) == 'ok':
            continue

        print(f'[{idx}/{len(codigos)}] {cod}...', end=' ', flush=True)
        data = obtener_info(s, cod)
        if not data:
            print('SIN DATOS')
            sin_datos.append(cod)
            progreso[cod] = 'sin_datos'
            time.sleep(PAUSA)
            continue

        cat_en = data.get('categoryName', '') or ''
        desc = TRADUCCION.get(cat_en, cat_en)
        if cat_en and cat_en not in TRADUCCION:
            sin_traducir.add(cat_en)
        cat = CAT_BUCKET.get(cat_en, 'General')
        if cat_en and cat_en not in CAT_BUCKET:
            sin_categoria.add(cat_en)

        props, ancho, alto, peso = construir_props(data)
        oe = construir_oe(data)
        compat, fabs = construir_compat(data)

        # Descargar imagenes: principal sin sufijo, extras _img2.._img6
        imagenes = data.get('imageList', []) or []
        nombre_principal = ''
        for n, url in enumerate(imagenes[:6], start=1):
            nombre = f'{cod}.jpg' if n == 1 else f'{cod}_img{n}.jpg'
            ruta = os.path.join(IMAGENES_DESTINO, nombre)
            if os.path.exists(ruta) or descargar_imagen(s, url, ruta):
                if n == 1:
                    nombre_principal = nombre

        nuevo = {
            'id': cod,
            'marca': 'BOGAP',
            'desc': desc,
            'cat': cat,
            'tipo': '',
            'lado': '',
            'props': props,
            'fabs': fabs,
            'oe': oe,
            'rel': '',
            'url': f'https://www.bogap.de/productsm/info.aspx?bogapNo={cod}',
            'img': f'Imagenes/Bogap/{nombre_principal}' if nombre_principal else '',
            'ancho': ancho,
            'alto': alto,
            'espesor': '',
            'avisador': '',
            'peso': peso,
            'ean': data.get('eanNo', ''),
            'compat': compat,
        }
        articulos.append(nuevo)
        agregados += 1
        progreso[cod] = 'ok'

        total_vehs = sum(len(v) for v in compat.values())
        print(f'OK -> {desc} [{cat}] | {len(compat)} marcas, {total_vehs} aplicaciones')

        if idx % 15 == 0:
            with open(ARTICULOS_JSON, 'w', encoding='utf-8') as f:
                json.dump(articulos, f, ensure_ascii=False, separators=(',', ':'))
            with open(PROGRESO, 'w', encoding='utf-8') as f:
                json.dump(progreso, f, ensure_ascii=False, indent=2)
            print(f'  [Guardado parcial {idx}/{len(codigos)}]')

        time.sleep(PAUSA)

    with open(ARTICULOS_JSON, 'w', encoding='utf-8') as f:
        json.dump(articulos, f, ensure_ascii=False, separators=(',', ':'))
    with open(PROGRESO, 'w', encoding='utf-8') as f:
        json.dump(progreso, f, ensure_ascii=False, indent=2)

    print()
    print('=' * 60)
    print(f'Agregados: {agregados}')
    print(f'Sin datos: {len(sin_datos)} {sin_datos}')
    if sin_traducir:
        print(f'\nCategorias SIN traduccion ({len(sin_traducir)}), quedaron en ingles:')
        for c in sorted(sin_traducir):
            print(' -', c)
    if sin_categoria:
        print(f'\nCategorias SIN bucket de cat, quedaron en "General" ({len(sin_categoria)}):')
        for c in sorted(sin_categoria):
            print(' -', c)
    print(f'\nTotal articulos.json: {len(articulos)}')


if __name__ == '__main__':
    main()
