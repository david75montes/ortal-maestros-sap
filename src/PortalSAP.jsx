import React, { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase.js";
import {
  Boxes, Tag, Lock, BookOpen, Database,
  Download, Upload, CheckCircle2, XCircle, AlertTriangle,
  ChevronRight, ChevronDown, X, ArrowRight, Sparkles, RotateCcw,
  Plus, Trash2, ShieldCheck, Pencil, Copy, ClipboardList, Clock, User, Inbox,
  Layers, ArrowLeftRight, Check, Save, Unlock,
  ShoppingCart, FileText, Truck, Percent, TicketPercent,
  Lightbulb, ArrowDown, ChevronUp, Menu, Home, Table2, Package,
  Utensils, Gift
} from "lucide-react";

const CENTROS = {"2003": "Pronto Pargua", "2549": "Pronto Rosario Norte", "2626": "Pronto Metro Quilicura", "2004": "Pronto Sierra Gorda", "2005": "Pronto Laguna Chicureo", "2006": "Pronto Manuel Rodriguez", "2615": "Pronto UAI Peñalolen", "2766": "Pronto Metro Pedro de Valdivia", "2778": "Pronto Carretera Longitudinal Austral", "2788": "Pronto Vicuña Mackenna", "2783": "Pronto Baquedano", "2100": "Pronto Antonio Rendic", "2101": "Pronto Pedro Aguirre Cerda", "2582": "Pronto Chañaral", "2102": "Pronto Diego Portales", "2796": "Pronto Ruta 27", "2103": "Punto Teniente Merino", "2510": "Pronto Antofagasta", "2722": "Punto Francisco Bilbao", "2105": "Pronto Vivar", "2106": "Pronto Ruta A-16", "2009": "Pronto Pozo Al Monte", "2107": "Punto Arturo Merino Correa", "2502": "Pronto Perez Zujovic", "2108": "Punto Granaderos", "2109": "Punto Panamericana Sur", "2110": "Pronto Luis Valente Rossi", "2111": "Pronto Angamos", "2112": "ProntoRuta 5 Norte", "2113": "Pronto Av Argentina I", "2758": "ProntoAntonio Rendic Esq.", "2114": "Pronto Barrio Industrial", "2115": "Pronto Pedro Prado", "2116": "Punto Abaroa", "2581": "Pronto Av Argentina II", "2118": "Punto 11 de Septiembre", "2119": "Pronto 18 de Septiembre", "2120": "Pronto Alto Hospicio", "2121": "Pronto Edmundo Perez Zujovic", "2122": "Punto Balmaceda", "2777": "Punto Diego Portales II", "2779": "Punto 27 de Abril", "2770": "Punto República de Croacia", "2514": "Pronto Iquique", "2123": "Pronto Santiago Arata", "2720": "Punto Anibal Pinto", "2124": "Punto Ruta 5 Sur", "2125": "Pronto Sotomayor", "2013": "Pronto San Rafael", "2012": "Pronto San Carlos", "2017": "Pronto Cosmito", "2540": "Pronto Paicaví", "2544": "Pronto Pedro De Valdivia", "2128": "Punto Rudecindo Ortega", "2737": "Punto Condell", "2053": "Pronto San José De La Mariquina", "2129": "Pronto Calle 8 Oriente", "2803": "Punto Camino Alerce", "2713": "Punto O'Higgins Chiguayante", "2131": "Punto Salvador Allende II", "2132": "Pronto Prat San Martin", "2133": "Punto Los Notros", "2134": "Punto Almirante La Torre", "2797": "Punto Pedro Aguirre Cerda II", "2135": "Pronto Vicente Mendez", "2056": "Pronto Puerto Varas", "2052": "Pronto Parral", "2055": "Pronto Trafun Poniente", "2059": "Pronto Victoria", "2063": "Pronto San Javier", "2136": "Pronto Futaleufu", "2705": "Pronto Rio Tranquilo", "2786": "Punto Maipu San Martin", "2137": "Pronto Ecuador O'Higgins", "2749": "Pronto Colin", "2138": "Punto Cardonal", "2757": "Punto Galvarino Riveros", "2139": "Pronto Prat II", "2140": "Pronto Kurt Moller", "2743": "Punto Lastarria", "2043": "Pronto Los Angeles Centro", "2721": "Punto O'Higgins III", "2142": "Punto Walker Martinez", "2143": "Punto Fresia Brasil", "2593": "Punto Rudecindo Ortega II", "2144": "Pronto Valdivia", "2145": "Punto San Martin Benavente", "2146": "Punto Av Mackenna", "2773": "Punto Arturo Alessandri", "2594": "Punto M. Rozas", "2719": "Pronto Sargento Aldea", "2148": "Pronto Balmaceda II", "2712": "Punto Bernardo O'Higgins", "2015": "Pronto Puerto Montt", "2794": "Pronto 2 Sur", "2150": "Punto Libertad", "2151": "Punto Picarte", "2152": "Punto Comercio", "2014": "Punto Jose Puchi", "2723": "Pronto Balmaceda", "2155": "Punto Mac-Iver", "2156": "Punto Caupolican", "2537": "Pronto Los Carrera", "2157": "Pronto Cristobal Colon", "2064": "Pronto Palomares", "2811": "Pronto Ladrilleros", "2159": "Punto Republica Victoria", "2048": "Pronto Chillán Oriente", "2160": "Pronto Presidente Frei", "2161": "Punto Geronimo De Alderete", "2162": "Punto Panamericana", "2163": "Pronto Manuel Rodriguez Chigu", "2164": "Pronto Pedro Aguirre Cerda", "2726": "Pronto 21 de Mayo", "2517": "Pronto Michimalonco", "2166": "Pronto 2 Norte", "2167": "Punto Picarte Simpson", "2168": "Punto Camino A Coñaripe", "2731": "Punto Los Carreras", "2590": "Punto Andres Bello", "2724": "Pronto Juan Antonio Rios", "2170": "Pronto Calle 5 Norte", "2171": "Punto Balmaceda III", "2065": "Pronto Cabrero", "2733": "Pronto Gabriela Mistral", "2173": "Punto Panamericana Sur II", "2030": "Pronto Horcones", "2706": "Pronto Alemania", "2175": "Pronto Lircay", "2176": "Punto Rene Soriano", "2177": "Pronto Vicuña Mackenna II", "2178": "Pronto Maraton", "2179": "Pronto Ruta S-30", "2180": "Pronto Camino Internacional", "2747": "Pronto La Junta", "2181": "Pronto Camino Internacional II", "2774": "Pronto Daniel Vera", "2182": "Pronto Av Argentina III", "2183": "Punto Vicente Perez Rosales", "2021": "Pronto Trafún Oriente", "2023": "Pronto Freire", "2020": "Pronto Loncoche", "2022": "Pronto Los Angeles Poniente", "2184": "Punto O'Higgins IV", "2185": "Punto Bulnes", "2186": "Punto Bulnes II", "2187": "Punto Bulnes III", "2188": "Punto Don Bosco", "2189": "Punto Ruta 5 Sur II", "2190": "Punto Los Rios", "2024": "Pronto Lautaro", "2191": "Pronto Los Carrera Lautaro", "2709": "Punto Castellon", "2018": "Pronto Eusebio Ibar", "2802": "Pronto Ogana", "2542": "Pronto Av Alemania", "2193": "Pronto Latorre", "2554": "Pronto Carlos Schorr", "2522": "Pronto Talcahuano", "2050": "Pronto Los Angeles Oriente", "2194": "Pronto Camino Coronel", "2195": "Punto O'Higgins Alemania", "2031": "Pronto Penco", "2800": "Punto Sargento Candelaria", "2196": "Punto Anibal Pinto II", "2197": "Punto Presidente Ibañez", "2730": "Punto Maipu 895", "2199": "Pronto Americo Vespucio", "2200": "Punto San Martin Uruguay", "2201": "Pronto Ercilla", "2202": "Punto Abate Molina", "2047": "Pronto Chillán Poniente", "2028": "Pronto Maule", "2203": "Punto Prat", "2204": "Pronto Calle B-20", "2205": "Pronto Los Carrera II", "2206": "Pronto Libertad", "2207": "Pronto Ruta 60-C", "2208": "Pronto Ruta F-30-E", "2209": "Pronto Lauro Barros", "2210": "Pronto Normandie", "2587": "Pronto Portales", "2211": "Pronto Colo Colo", "2051": "Pronto Hijuelas", "2784": "Punto Los Loros", "2212": "Pronto Sta. Teresita Peñablanca", "2213": "Pronto Agua Santa", "2214": "Pronto Valparaiso", "2543": "Pronto Bosques de Montemar", "2215": "Pronto Bernardo O'Higgins II", "2742": "Pronto Esmeralda Santa Teresa", "2217": "Pronto Humeres", "2058": "Pronto Palo Colorado", "2218": "Punto O'Higgins V", "2219": "Pronto Palmira Romano Norte", "2060": "Pronto Marbella", "2220": "Pronto Marga Marga", "2221": "Pronto Teniente Cruz Martinez", "2519": "Pronto Las Salinas", "2746": "Pronto Carretera Panamericana Norte", "2222": "Pronto Freire Echeverria", "2223": "Pronto Freire Gomez Carreño", "2703": "Pronto Jose Joaquin Perez", "2225": "Pronto Constitucion", "2069": "Pronto Placilla", "2226": "Punto Panamericana Norte II", "2567": "Punto Brasil I", "2228": "Punto Ignacio De La Carrera", "2229": "Punto Panamericana Norte Fco. Aguirre", "2230": "Pronto Balmaceda IV", "2231": "Punto Ariztia", "2057": "Pronto Socos", "2232": "Pronto El Peral", "2233": "Pronto Lusitania", "2234": "Pronto Los Carrera III", "2235": "Pronto Irarrazaval O'Higgins", "2236": "Pronto Borgoño", "2237": "Pronto Playa Ancha", "2577": "Pronto Alessandri", "2239": "Pronto Irarrazaval", "2240": "Pronto Las Delicias Sargento Aldea", "2241": "Pronto Balmaceda Paradero", "2242": "Pronto Vicuña Mackenna III", "2243": "Pronto Chacabuco General Cruz", "2244": "Punto Alessandri Norte", "2245": "Pronto San Martin 2510", "2246": "Pronto El Arrayan", "2044": "Pronto Los Vilos", "2247": "Pronto Marga Marga II", "2716": "Pronto Camino Internacional Oriente", "2805": "Pronto Rafael Ariztia", "2249": "Pronto Matta Cerro Placeres", "2250": "Pronto Colo Colo", "2251": "Pronto Circunvalacion Ariztia", "2252": "Pronto Tocornal", "2253": "Pronto Adolfo Eastman", "2254": "Punto Merino Jarpa", "2255": "Pronto La Cantera I", "2256": "Pronto Ruta F30", "2257": "Pronto Camino Troncal", "2258": "Pronto Curauma Sur", "2259": "Pronto Isidoro Dubournais", "2260": "Pronto Jose M. Balmaceda", "2708": "Pronto Emilio Valle", "2262": "Pronto Juan Rusque", "2263": "Pronto Palmira Romano Sur", "2771": "Pronto Gabriel Gonzalez Videla", "2732": "Pronto Av Argentina , Las Juntas", "2025": "Pronto Los Vilos Oriente", "2265": "Pronto La Cantera II", "2266": "Pronto Calle Larga", "2267": "Punto Socos Victoria", "2268": "Pronto Tarcicio Valderrama", "2269": "Pronto Januario Ovalle", "2270": "Punto Carretera Panamericana III", "2096": "Pronto Copiapo Km 811", "2271": "Punto Copayapu", "2089": "Pronto Copiapo Km 838", "2272": "Pronto El Islon", "2273": "Pronto Alessandri II", "2595": "Pronto 4 Esquinas", "2534": "Pronto Parcela 20", "2275": "Pronto Ambrosio Ohiggins", "2045": "Pronto Coquimbo", "2276": "Pronto Camino Internacional II", "2277": "Pronto Los Carrera N° 01050", "2810": "Pronto Nicaragua", "2041": "Pronto Llay Llay", "2278": "Punto Camino Tuqui", "2040": "Pronto La Serena", "2046": "Pronto Tabolango", "2545": "Pronto Troncal Sur", "2279": "Pronto Los Copihues", "2280": "Pronto Juan Bautista Alberdi", "2011": "Pronto Puente Ruta 78", "2536": "Pronto Isidora", "2612": "Pronto Plaza Peru", "2614": "Pronto UDD Concepción", "2616": "Pronto UAI Viña Del Mar", "2618": "Pronto Metro Vicuña Mackenna", "2619": "Pronto USS Cede Los Leones", "2621": "Pronto Casino UNAB", "2622": "Pronto Casino UAI Peñalolen", "2624": "Pronto DUOC Vina Del Mar", "2775": "Pronto Lyon", "2761": "Pronto Plaza Sucre", "2763": "Pronto Plaza de Armas Temuco", "2764": "Pronto Mall Plaza El Trebol", "2765": "Pronto Metro Cal Y Canto", "2767": "Pronto Valparaiso", "2768": "Pronto Bellavista Pio Nono", "2780": "Pronto Luis Thayer Ojeda", "2789": "Pronto UAI Viña Del Mar Casino", "2790": "Pronto UAI Peñalolén Casino Edif. A", "2791": "Pronto UAI Peñalolén Casino Edif. E", "2792": "Pronto UAI Peñaloles Cafetería Edif. C", "3014": "Pronto CIUC Piso 1", "2762": "Pronto Huérfanos 815", "8201": "Pronto Metro U De Chile", "2500": "Pronto Pedro Fontova", "2532": "Pronto Chamisero", "2010": "Pronto Nos", "2281": "Pronto Balmaceda Malloco", "2282": "Pronto General San Martin", "2283": "Pronto Alcalde Lopez", "2617": "Pronto Macul", "2782": "Pronto Americo Vespucio II", "2284": "Punto Errazuriz", "2285": "Pronto Pedro Aguirre Cerda II", "2062": "Pronto Ruta 78 Poniente", "2061": "Pronto Ruta 78 Oriente", "2287": "Pronto Bilbao", "2288": "Pronto Vitacura", "2289": "Pronto Macul", "2290": "Punto San Martin N° 401", "2291": "Pronto Ecuador", "2292": "Pronto General Prieto", "2293": "Pronto Balmaceda II", "2509": "Pronto Vitacura 5579", "2511": "Pronto Av La Florida", "2294": "Pronto Blanco Encalada", "2531": "Pronto Cantagallo", "2296": "Pronto Pedro Aguirre Cerda III", "2812": "Pronto Ruta G16 Arco Iris", "2298": "Pronto Irarrazaval II", "2299": "Punto Chicureo", "2801": "Pronto Panamericana Norte", "2300": "Pronto San Ramon", "2539": "Pronto Pajaritos 5200", "2301": "Pronto Irarrazaval III", "2054": "Pronto San Fernando", "2302": "Pronto San Pablo Bismark", "2541": "Pronto Costanera E0", "2303": "Pronto Bernardo O'Higgins I", "2578": "Pronto Camino Melipilla II", "2305": "Pronto Camino Nos", "2306": "Pronto Eliodoro Yañez", "2307": "Pronto Concha Y Toro I", "2308": "Pronto Oriental", "2813": "Pronto El Rosal", "2310": "Pronto Salvador", "2311": "Pronto Membrillar", "2312": "Pronto Lota", "2513": "Pronto Portal La Dehesa", "2019": "Pronto Nos Km 27", "2313": "Pronto Vicuña Mackenna IV", "2717": "Pronto Colon", "2754": "Pronto Pedro De Valdivia", "2315": "Pronto Ossa", "2515": "Pronto Principe De Gales", "2704": "Pronto Santa Rosa II", "2317": "Punto Gran Avenida", "2318": "Pronto San Pablo II", "2586": "Pronto Mapocho Brasil", "2320": "Punto Jose Miguel Carrera", "2321": "Pronto Americo Vespucio III", "2707": "Pronto General San Martin", "2323": "Pronto Departamental", "2324": "Pronto Av San Juan", "2325": "Pronto España", "2326": "Pronto O'Higgins Lote B-5A", "2327": "Pronto Arturo Pratt", "2002": "Pronto Sagrada Familia", "2328": "Pronto Circunvalacion Norte", "2714": "Pronto Alberto Llona", "2330": "Pronto Concha Y Toro II", "2331": "Pronto Recoleta", "2769": "Pronto San Pablo Brasil", "2575": "Pronto Lia Aguirre", "2333": "Pronto Bernardo O'Higgins III", "2334": "Punto Salvador Gutierrez", "2335": "Pronto Longitudinal Sur", "2336": "Punto Carrascal", "2337": "Pronto El Valle", "2338": "Punto Mapocho", "2339": "Pronto Manso De Velasco", "2340": "Pronto Carretera El Cobre", "2772": "Punto Irarrazabal II", "2341": "Pronto Recoleta", "2342": "Pronto Vespucio", "2343": "Pronto 18 De Septiembre", "2344": "Pronto 5 De Abril", "2345": "Pronto Libertador Bernardo O'Higgins", "2346": "Pronto Walker Martinez", "2793": "Pronto Tobalaba", "2347": "Pronto Bernardo O'Higgins II", "2503": "Pronto Pajaritos 3333", "2049": "Pronto Ruta 68", "2348": "Pronto Alameda", "2349": "Pronto Riesgo", "2566": "Pronto Ignacio Carrera Pinto", "2351": "Punto Orlandi", "2088": "Pronto Chimbarongo", "2352": "Pronto Longitudinal Sur II", "2715": "Pronto Carmen Camilo Henriquez", "2745": "Pronto Independencia", "2354": "Pronto Longitudinal Sur III", "2355": "Punto Argomedo", "2356": "Pronto Viel", "2785": "Pronto Bustamante", "2357": "Pronto Freire II", "2358": "Pronto Camino Melipilla II", "2359": "Punto San Martin M. Solis", "2360": "Pronto Bascuñan Guerrero", "2807": "Pronto Manuel Montt", "2361": "Pronto Cachapoal", "2565": "Pronto Manquehue", "2362": "Pronto Departamental II", "2363": "Pronto Las Condes", "2729": "Pronto Vicuña Mackenna V", "2365": "Pronto Los Leones", "2727": "Pronto Carmen", "2367": "Pronto Consistorial", "2368": "Pronto Calera De Tango", "2711": "Pronto Arturo Prat II", "2370": "Punto San Eugenio", "2806": "Pronto Camino Lo Sierra", "2371": "Punto Jose Pedro Alessandri", "2372": "Pronto Francisco Bilbao II", "2804": "Pronto Diagonal Oriente", "2373": "Punto Alberto Edwards", "2718": "Pronto Camino El Alba", "2776": "Pronto Ortuzar", "2374": "Pronto Trinidad", "2728": "Pronto Providencia", "2026": "Pronto San Fco Mostazal Poniente", "2027": "Pronto San Fco Mostazal Oriente", "2376": "Pronto San Jose", "2548": "Pronto Trapenses", "2377": "Pronto Camino Lonquen", "2505": "Pronto Lord Cochrane", "2378": "Pronto Americo Vespucio", "2787": "Pronto Virginia Subercaseaux", "2379": "Punto Jose Joaquin Perez II", "2556": "Pronto Vitacura 4207", "2530": "Pronto Vicuña Mackenna 1990", "2380": "Pronto Bascuñan Guerrero II", "2381": "Pronto Camino Lo Ovalle", "2576": "Pronto Apoquindo", "2504": "Pronto Vicuña Mackenna 5700", "2506": "Pronto La Dehesa", "2016": "Pronto Lampa", "2525": "Pronto Costanera E1", "2526": "Pronto Costanera E2", "2029": "Pronto Costanera E6", "2507": "Pronto San Pablo", "2568": "Pronto Carmen III", "2512": "Pronto Santa Maria", "2086": "Pronto Requinoa", "2087": "Pronto Romeral", "2384": "Pronto 21 De Mayo", "2385": "Pronto Mexico El Peñon", "2508": "Pronto Las Condes 10912", "2599": "Pronto Americo Vespucio IV", "2387": "Pronto Recreo", "2589": "Pronto Gabriela", "2518": "Pronto C Henriquez", "2521": "Pronto Los Libertadores", "2388": "Pronto Leonardo Murialdo", "2781": "Pronto Comercio", "2389": "Pronto Concha Y Toro III", "2390": "Pronto Santa Rosa III", "2538": "Pronto Tobalaba", "2795": "Pronto Austral", "2391": "Punto Henriquez", "2392": "Pronto Manzano", "2799": "Pronto Ruta 215", "2798": "Pronto Matta III", "2007": "Pronto Panamericana Norte II", "2633": "Pronto Irarrazaval Brown Sur", "2637": "Pronto Antonio Bellet", "2394": "Pronto General Velasquez", "2630": "Pronto Colo Colo", "2808": "Pronto Padre Las Casas", "2809": "Pronto Tijerales", "2638": "Pronto San Carlos de Apoquindo", "2396": "Pronto Mejillones", "2634": "Pronto Holanda", "2648": "Pronto Mall Plaza Alameda", "2631": "Pronto Agustinas San Antonio", "2397": "Pronto Quillon", "2641": "Pronto Antofagasta Arauco Express", "2645": "Pronto Agustinas Morande", "2405": "Pronto San Diego", "2643": "Pronto Padre Mariano", "2104": "Pronto Tinuche", "2639": "Pronto La Florida El Membrillar", "2651": "Sbarro Mall Plaza Norte", "2650": "Sbarro Mall Plaza Egaña", "2635": "Pronto Bulnes-Portales Temuco", "2652": "Sbarro Mall Costanera Center", "2552": "Pronto Quilin", "2066": "Pronto Monterilla", "2036": "Pronto Vallenar", "2130": "Pronto Villarrica", "2141": "Pronto Colegio Ingles talca", "2147": "Pronto Huamachuco", "2117": "Pronto Los Carrera", "2149": "Pronto Lorca", "2127": "Pronto San Carlos Sur", "2814": "Pronto Pichilemu Comercio", "2153": "Pronto Laura Pizarro", "7002": "Juan Valdez Providencia", "7003": "Juan Valdez M Pza Dominicos", "7004": "Juan Valdez MUT", "7005": "Juan Valdez M Plaza La Serena", "7006": "Juan Valdez M Plaza Los Ríos", "7007": "Juan Valdez Clínica Las Condes", "7008": "Juan Valdez M Plaza Trebol", "7009": "Juan Valdez M Portal Osorno", "7010": "Juan Valdez M Costanera PM", "7011": "Juan Valdez Aires Vespucio", "7012": "Juan Valdez M Arauco Maipu", "7013": "Juan Valdez M Pza Egana Corn", "7014": "Juan Valdez Rosario Norte", "7015": "Juan Valdez M Marina Arauco", "7016": "Juan Valdez Catedral", "7017": "Juan Valdez M Independencia", "7018": "Juan Valdez Open Kennedy", "7019": "Juan Valdez Apto Espigon C", "7020": "Juan Valdez Apto Espigon E", "7021": "Juan Valdez M P Arauco Corn", "7022": "Juan Valdez M Parque Arauco", "7023": "Juan Valdez M Costanera Ctr", "7024": "Juan Valdez M Plaza Egana", "7025": "Juan Valdez Open Pza Rancagua", "7026": "Juan Valdez M Alto Las Condes", "7027": "Juan Valdez M Pza Antofagasta", "7028": "Juan Valdez M Plaza Tobalaba", "7029": "Juan Valdez Lider Puente Nuevo", "9005": "Juan Valdez Darkstore Wallmart", "9008": "Juan Valdez Darkstore Justo", "9003": "Centro Distribucion Bluex", "9002": "Centro Distribucion 3PL"};
const SKUS = {"11": "BULTO", "32": "COCA-COLA ZERO 250C (TUR BUS)", "51": "MAT_MOD COMP AROS DE CEBOLLA", "71": "DESPACHO", "82": "AIRE ACOND 60.000 MIDEA", "83": "AIRE ACOND 240.000 DAIKIN", "84": "CAMARA FRIO 6.05 REFMAR", "10000": "MAT_MOD INSUMO MATERIAS PRIMAS", "10002": "QUESO DE CABRA", "10003": "TRUTRO DE POLLO ENTERO FRESCO", "10004": "POSTA ROSADA (GANSO)", "10005": "LECHE ENTERA", "10006": "QUESO MOZZARELLA", "10007": "QUESO CREMA", "10008": "QUESO REGGIANITO   (PARMESANO)", "10009": "JAMON DE PAVO LAMINADO", "10010": "PASTRAMI LAMINADO EE", "10011": "LONGANIZA CORTADA (PIZZAS)", "10012": "JAMON LAMINADO PIZZA", "10013": "HAMBURGUESA 70 GR VACUNO", "10014": "ARROLLADO HUASO", "10015": "PECHUGA POLLO DESHUESADA FRESCA", "10016": "FILETE DE POLLO GRILLADO COCIDO CONG", "10017": "ALBONDIGAS (CRUDA)", "10018": "JAMON SANDWICH LAMINADO", "10019": "CARNE MECHADA", "10020": "JAMON PRAGA GRADO 2 (PIEZA)", "10021": "SALAME TIPO ITALIANO (PIEZA)  (ST)", "10022": "TOCINO (PANCETA AHUMADA)", "10023": "PAN FRICA 10 CON SESAMO", "10024": "NO UTILIZAR", "10025": "HUEVO DE CODORNIZ 1×24 UN", "10026": "PAN CIABATTA GRANDE", "10027": "PAN MOLDE MIGA S/CORTEZA", "10028": "PAN MOLDE BLANCO 900G", "10029": "PAN MOLDE INTEGRAL CON BORDE", "10030": "PAN MOLDE AMARILLO", "10031": "PAN MOLDE SOUR", "10032": "PAN MOLDE ACEITUNA", "10033": "PAN MOLDE MIGA GIGANTE BLANCO (BOL)", "10034": "PAN MOLDE MIGA GIGANTE INTEGRAL (BOL)", "10035": "PAN CROISSANT 80 GRMS", "10036": "PAN PANINI", "10037": "PAN FRANCES (BAGUETTE)", "10038": "PAN CROCATA PRECOCIDA", "10039": "MARGARINA CONDESA BATERKRIM VEGETAL 20X1", "10040": "MASA PIZZA INDIVIDUAL", "10041": "SALMON AHUMADO CONGELADO", "10042": "CALAMAR CONGELADO", "10043": "(NO UTILIZAR) CAMARON CONGELADO", "10044": "ESPARRAGO CONGELADO", "10045": "ACEITUNAS NEGRAS RODAJAS", "10046": "HUEVO DURO PASTEURIZADO 1×30 UN", "10047": "TOMATE (ENTERO)", "10048": "PIMENTON JULIANA ROJO", "10049": "QUESO CHACRA (FRESCO)", "10050": "POROTOS VERDES CORTADOS", "10051": "MANZANA VERDES", "10052": "FRUTILLAS", "10053": "TUNAS", "10054": "ZANAHORIA PELADA", "10055": "UVA", "10056": "NARANJA", "10057": "PEPINO ENSALADA", "10058": "PIMENTON ROJO (ENTERO)", "10059": "PIMENTON VERDE (ENTERO)", "10060": "PIMENTON MECHADA ROJO", "10061": "PIMENTON MECHADA VERDE", "10062": "BERROS HIDROPONICOS", "10063": "LENTEJAS", "10064": "PALTA HASS", "10065": "PAPA FRITA HILO", "10066": "LECHUGA ESPANOLA", "10067": "LECHUGA LOLO ROSE", "10068": "PAPA CUBO", "10069": "HUESILLOS", "10070": "CIBOULETTE", "10071": "PREMEZCLA MASA PIZZA NEUTRA 1*25", "10072": "CEBOLLA CUBO", "10073": "SEMILLA DE SESAMO", "10074": "CEBOLLA PELADA", "10075": "CEBOLLA PLUMA", "10076": "SANDIA", "10077": "PINA", "10078": "MELON", "10079": "MOTE CON HUESILLOS", "10080": "ZANAHORIA DESPUNTE", "10081": "PEPINILLOS", "10082": "ANCHOAS", "10083": "HUEVO BLANCO", "10084": "PEREJIL", "10085": "NORY", "10086": "ESENCIA LUCUMA", "10087": "SEMILLA DE AMAPOLA KG", "10088": "CEREZAS CONFITADAS", "10089": "WHASABI", "10090": "AJI VERDE", "10091": "AJO PELADO", "10092": "ALCAPARRA", "10093": "ACELGA", "10094": "ZAPALLO CAMOTE", "10095": "ZAPALLO ITALIANO", "10096": "CEBOLLA PERLA", "10097": "ESPINACA PROCESADA", "10098": "KIWI", "10099": "QUESO PHILADELPHIA", "10100": "(NO UTILIZAR) CONAC", "10101": "MARGARINA 1×12 KG (NO UTILIZAR)", "10102": "LECHUGA MARINA", "10103": "PAPAS PELADAS (ENTERAS)", "10104": "CEBOLLA PLUMA DOLE", "10105": "LECHUGA ESCAROLA MAC DOLE 1 KG", "10106": "LECHUGA ESCAROLA SUPER", "10107": "ENSALADA 4 ESTACIONES", "10108": "MIX 3", "10109": "QUESO GAUDA ENTERO (PIEZA)", "10110": "QUESO GAUDA GRANULADO", "10111": "PLATEADA (CRUDA)", "10112": "POSTA PICADA 1 POR 1", "10113": "LOMO DE CERDO (CRUDO)", "10114": "MOSTAZA", "10115": "PURE DE PAPAS", "10116": "CALDO DE GALLINA 6*1 KG", "10117": "CONCENTRADO DE CARNE KG", "10118": "SALSA DEMI GLACE 6*1 KG", "10119": "CONCENTRADO DE TOMATE", "10120": "JAMON SERRANO (LAMINADO)", "10121": "JAMON IBERICO", "10122": "MAYONESA", "10123": "CHAMPINION LAMINADO EN CONSERVA", "10124": "ATUN EN ACEITE POUCH 1KG", "10125": "ENELDO", "10126": "PALMITOS EN CONSERVA (ENTERO)", "10127": "PINAS CUBO EN CONSERVA", "10128": "CEREZAS EN CONSERVA (DREN. 1,6 KG)", "10129": "PIMENTON EN CONSERVA", "10130": "FETUCCINI AL HUEVO (CRUDO)", "10131": "(NO UTILIZAR) SALSA PARA PIZZA", "10132": "PASTA DE CHOCLO", "10133": "FONDOS ALCACHOFA EN CONSERVA NO OCUPAR", "10134": "AZUCAR CASINO", "10135": "ACEITE VEGETAL 5LT", "10136": "VINO TINTO", "10137": "FRUTA EN CONSERVA DURAZNO KG", "10138": "NUECES MARIPOSA", "10139": "SUCEDANEO JUGO DE LIMON", "10140": "SAL DE MESA", "10141": "PASTA TIPO ESPIRAL KG", "10142": "OREGANO ENTERO", "10143": "FLAN DE VAINILLA", "10144": "REFRESCO SABORES", "10145": "MOUSSE CHOCOLATE (POLVO)", "10146": "CHANCACA", "10147": "(NO UTILI) ARROZ LARG G-1 PREGRANEADO", "10148": "ACEITUNA NEGRA DESCAROZADA", "10149": "ARVEJAS CONGELADAS", "10150": "CHOCLO", "10151": "AGA PACK 2% CO2", "10152": "AGA PAC 3.1", "10153": "ACEITE DE OLIVA", "10154": "AJI DE COLOR", "10155": "CARNE MOLIDA ESPECIAL", "10156": "COMINO", "10157": "LASAGNA 1*400GR", "10158": "LAUREL DESHIDRATADO KG", "10159": "ALMIDON DE MAIZ MAICENA", "10160": "(NO UTILIZAR) NUEZ MOSCADA", "10161": "PASAS", "10162": "POROTOS BURRO", "10163": "SALSA BLANCA (POLVO)", "10164": "SALSA DE SOYA", "10165": "JIBIA CONGELADA", "10166": "ACIDO CITRICO", "10167": "BENZOATO DE SODIO", "10168": "SORBATO DE POTACIO", "10169": "MANI", "10170": "PERA (NO UTILIZAR)", "10171": "CILANTRO", "10172": "MANJAR 4*4,5KG", "10173": "AGUA ACQUA SOFT GAS 500CC", "10174": "MASA DE BURRITO", "10175": "CURRY", "10176": "KETCHUP", "10177": "MERQUEN", "10178": "CLAVOS DE OLOR", "10179": "REPOLLO MORADO (NO USAR)", "10180": "HUEVO LIQUIDO GALLINAS LIBRES", "10181": "CHAMPINON NATURAL", "10182": "CREMA FRESCA LITRO", "10183": "PIMIENTA BLANCA MOLIDA", "10184": "JENJIBRE", "10185": "ESCALOPAS CORTE JULIANA (CRUDO)", "10186": "MIX SALAD DOLE", "10187": "ARROZ G2 LARGO Y ANCHO", "10188": "CREMA VEGETAL", "10189": "BROTE DE ALFALFA", "10190": "EMPANADA DE PINO 220 GR", "10191": "BICARBONATO", "10192": "JALEA PINA", "10193": "PEPINO FRUTA", "10194": "PIMENTON CONSERVA CUBITOS", "10195": "PANACHE DE VERDURAS", "10196": "MENTA", "10197": "ZANAHORIA HILO", "10198": "(NO UTILIZAR) DIENTES DE DRAGON", "10199": "MERMELADA DE FRAMBUESA", "10200": "MERMELADA DE MORA", "10201": "MERMELADA DE DURAZNO", "10202": "PULPA DE CERDO", "10203": "MOUSSE MANJAR", "10204": "CAMARON CONGELADO (DESCALIB)", "10205": "CHULETAS DE CERDO  150 GR (CRUDA)", "10206": "NO UTILIZAR (PULPA DE CERDO)", "10207": "DIENTE DE DRAGON", "10208": "(NO UTILI)LECHUGA  COSTINA PROCESADA", "10209": "CEBOLLIN PICADO", "10210": "CHOCLO ENANO", "10211": "QUESO GRUYERE", "10212": "POROTOS NEGROS", "10213": "PAN MARRAQUETA", "10214": "MASA DE BURRITO 18 CM", "10215": "NACHOS", "10216": "(NO UTILIZAR) BAGEL CHIPS", "10217": "(NO UTILIZAR) GRISSINNI", "10218": "ALBAHACA (NO OCUPAR)", "11005": "(NO UTLILIZAR) GAS AGA 2% CO2", "11010": "GAS AGA 3.1", "11011": "PAPAYAS", "11013": "POLLO CUBO (COCIDO)", "11014": "RIGATONI", "11015": "PAPAS GAJOS SKIN ON", "11016": "SALSA GAMBERO DE MARE", "11017": "TRUTRO DESHUESADO FRESCO CON PIEL", "11021": "POLLO ASADO ENTERO", "11022": "MOUSSE MARACUYA", "11023": "(NO UTILIZAR) MOUSSE CHOCOLATE BLANCO", "11024": "(NO UTILIZAR) TIRAMISU", "11031": "NO UTILIZAR", "11032": "LECHE DE COCO", "11041": "RAVIOLI", "11042": "PATE 1×125 GR.", "11043": "VIENESAS", "11044": "ZANAHORIA VICHY", "11051": "COSTILLAR DE CERDO", "11052": "AJI CREMA", "11053": "EMPANADA DE PINO COCKTAIL (ROH)", "11054": "PASTA TIPO ESPAGUETI KG", "11061": "CHURRASCO SANDWICH 70 GR.", "11071": "PIMENTON PAISANO", "11072": "ZANAHORIA CUBO (NO UTILIZAR)", "11073": "TOMATE COCKTAIL (NO UTILIZAR)", "11074": "HUACHALOMO PICADO", "11081": "SALMON AHUMADO INTERFOLIADO", "11084": "POLLO ASADO (UNIDAD)", "11092": "NO UTILIZAR", "11101": "JAMON PLANCHADO CUBO", "11102": "JAMON PIERNA GRADO 1", "11111": "PALMITO EN RODAJA (MANZANA)  (ST)", "11112": "PALMITO MOLIDO", "11121": "FILETILLO DE POLLO (PAGODA CRUDO)", "11122": "PECHUGA DE PAVO CONGELADA", "11132": "TE CLUB ROYAL CEYLAN", "11133": "CAFE NESCAFE 12*170 GRMS.", "11141": "CREMA VERDURAS", "11142": "CREMA ESPARRAGOS", "11143": "CREMA POLLO", "11144": "GALLETAS", "11151": "MASA PIZZA FAMILIAR", "11154": "PALTA TROZADA", "11155": "SACHET CESAR", "11156": "SACHET VINAGRETA", "11161": "(NO UTILIZAR) ALBUMINA PASTEURIZADA", "11162": "ALMENDRA COMERCIAL ENTERA", "11163": "ALMENDRA PICADA", "11165": "ARANDANO IQF", "11166": "AVENA MACHACADA", "11167": "AZUCAR FLOR", "11168": "AZUCAR RUBIA", "11169": "LADY MIROIR (FRUTGEL) 1*5KG", "11170": "CANELA MOLIDA", "11171": "CASTANAS PURE.", "11172": "CHOCOLATE GOTA FLEISCHMANN S/AMARGO", "11173": "COBERLUX (DEBELIS BITTER) 5×10", "11174": "CHOCOLATE EN POLVO.", "11175": "COBERTURA ALFAJOR.", "11176": "CARIBE GOTITAS SEMIAMARGO", "11177": "COBERTURA BG 1×5", "11178": "COCO RALLADO", "11179": "COLDFIL FRAMBUESA.", "11180": "ESCENCIA DE NARANJA", "11181": "CREMA SACHET.", "11182": "CREMA VEGETAL WHIPPAK.", "11183": "CREMIVIT 12 × 400 GRM (4,8KG)", "11184": "(NO UTILIZAR)HARINA", "11185": "MANJAR RECETA CASERA", "11186": "DULCE DE MORA.", "11187": "DULCE MEMBRILLO. (NO OCUPAR)", "11189": "ESCENCIA DE ALMENDRA", "11190": "FRAMBUESA CRUMBLE", "11191": "COLORANTE DE LUCUMA", "11193": "FRUGOLIM LIMON.", "11194": "FRUIT FILLING MANZANA.", "11195": "FRUTA CONFITADA.", "11196": "GELATINA GRANEL.", "11197": "MERENGUE 10X500", "11198": "PREMEZCLA INTEGRAL ESPECIAL", "11199": "PREMEZCLA BRAZO REINA 1×25", "11201": "ESENCIA DE VAINILLA", "11203": "YEMA LIQUIDA", "11204": "PREMEZCLA GALLETA 1× 10  KG", "11205": "PREMEZCLA MASA PIZZA SABORIZADA 1*25", "11206": "PREMEZCLA BROWNIE", "11207": "QUESO RICOTTA QUILLAYES", "11209": "(NO UTIZAR) SEMOLA INSTITUCIONAL", "11210": "TEGRAL BERLINA", "11211": "TEGRAL FACTURA", "11212": "MERMELADA GUINDA", "11213": "TEGRAL SATIN CREAM", "11214": "TEGRAL SOFTER HAMBURGER", "11215": "LEVADURA, OKEDO 20×500GR", "11216": "TEGRAL PAN ITALIANO", "11217": "TEGRAL BROWNIE", "11218": "MERMELADA LUCUMA", "11219": "MERMELADA DAMASCO", "11220": "HARINA DE ALMENDRAS", "11221": "SALSA PARA PIZZA 10*1 KG", "11222": "HARINA DE ARROZ", "11223": "NARANJAS CONFITADAS CUBO", "11224": "HARINA DE SOYA", "11225": "HARINA ORO", "11226": "HARINA PLATA", "11227": "MARGARINA HORNITO HORNEO 20X1", "11229": "JUGO NARANJA CONGELADO", "11230": "CHOCOLATE BELCOLADE BLANCO", "11232": "HUEVO EN POLVO", "11233": "NUEZ CUARTILLO", "11234": "LECHE CONDENSADA 4 × 4,5 KG", "11235": "LECHE EN POLVO", "11236": "LECHE EVAPORADA", "11237": "VAINILLINA", "11238": "(NO UTILIZAR) LEVADURA", "11239": "MARGARINA HORNITO HOJA 20X1", "11240": "MANI GRANILLO", "11241": "MANZANA DESHIDRATADAS", "11242": "ALMENDRA SIN PIEL", "11243": "NUTELLA 15*350 GRMS", "11244": "MASA FILLO", "11245": "(NO UTILIZAR) PAPAYA STICK BASTONES", "11246": "PAPAYAS  GALON 3.2 KG", "11247": "PASAS  RUBIAS", "11249": "PASAS CORINTO", "11250": "MANTEQUILLA SIN SAL ANCHOR (NO UTILIZAR)", "11251": "CHOCOLATE FLEISCHNAMM LECHE", "11252": "PASTA MANI", "11253": "PISTACHOS", "11254": "POLVOS DE HORNEAR 10×500 GR.", "11255": "PREMEZCLA BISCOCHO CHOCOLATE 1×25", "11256": "PREMEZCLA BISCOCHO INDUSTRIAL 1×30", "11257": "DATILES SIN CAROZO", "11258": "(NO UTILIZAR) JARABE DE REMOJO", "11259": "(NO UTILIZAR) BRAZO DE REINA VAINILLA", "11260": "(NO UTILIZAR) PLANCHA DE BIZCOCHO CHOCOL", "11261": "(NO UTILIZAR) PLANCHA DE BIZCOCHO VAINI", "11262": "(NO UTILIZAR) PRE TORTA ALFAJOR 15PP", "11263": "(NO UTILIZAR) PLANCHA DE HOJA (MP)", "11264": "(NO UTILIZAR)PLANCHA PANQUEQUE CHOC.(MP)", "11265": "(NO UTILIZAR)PLANCHA PANQUEQUE VAIN(MP)", "11271": "(NO UTILIZAR) ALBONDIGAS", "11281": "ACEITUNA VERDE RODAJA", "11291": "POSTA CUBO 2 X 2 (ESTOFADO Y GOULASH)", "11292": "MANZANA ROJA", "11301": "PEPPERONI", "11311": "POLLO COCIDO C&C", "11312": "CHORIZO SARTA (PIZZAS)", "11321": "YOGURT MORA", "11322": "YOGURT PINA", "11323": "YOGURT CHIRIMOYA", "11324": "YOGURT NATURAL 1*900ML", "11325": "MERMELADA (FRUIT FILLING PINA)", "11326": "JUGO NARANJA UHT", "11327": "JUGO EN POLVO KG", "11328": "QUESO ROQUEFORT", "11329": "PLATANO", "11330": "PERA", "11331": "BISCOCHO HOJALDRE LARGUITA (50G) PRONTO", "11332": "PALMERA HOJALDRE (150GR) PRONTO", "11342": "BISCOCHO REDONDO VAINILLA", "11343": "DESMOLDANTE (BLOQUEO AA)", "11344": "MERENGUE TORTA FRAMBUESA", "11345": "BIZCOCHO TORTA TRUFA", "11351": "PREMEZCLA PIZZA INTEGRAL", "11352": "CHORICILLO CORTADO", "11361": "CEBOLLA CARAMELIZADA", "11362": "POLLO ENTERO", "11371": "POROTO VERDE C/FRANCES CONGELADO 4X2 KG", "11372": "ESPINACA CONGELADA KG", "11373": "JAMON CORTE JULIANA", "11374": "PANO MULTIUSO AMARILLO", "11381": "BISCOCHO HOJALDRE LARGUITA (150G) PRONTO", "11382": "MAYONESA HELLMANS DELI", "11383": "PECHUGA POLLO CUARTO", "11384": "SALSA ITALIANA BLS 10X1KG", "11385": "MANJAR PASTELERO", "11391": "COCA-COLA LATA DE 350 CC", "11392": "ALMENDRA FILETILLO", "11393": "PIMIENTA NEGRA MOLIDA", "11401": "AJO EN POLVO", "11402": "QUESO PARMESANO ESCAMAS  (NO UTILIZAR)", "11403": "JUGO DURAZNO 200 CC", "11411": "LOMITO CERDO COCIDO (LAMINADO)", "11412": "ALBAHACA DESIDRATADA", "11413": "CIBOULETTE DESHIDRATADO", "11414": "OREGANO MOLIDO", "11415": "PEREJIL DESHIDRATADO KG", "11421": "DURAZNO", "11422": "PAPA COCIDA", "11423": "ZANAHORIAS CUBO", "11431": "MASA PIZZA FAMILIAR A LA PIEDRA", "11451": "REPOLLO MORADO PICADO", "11452": "APIO JULIANA", "11453": "COLIFLOR PROCESADA", "11454": "REPOLLO BLANCO PICADO", "11455": "BETARRAGA CUBO", "11456": "MEDALLON DE PAVO", "11461": "JAMON PRAGA", "11462": "PULPA DE CERDO PORCIONADA", "11472": "VINAGRE ROSADO 24X500CC", "11481": "MEDALLON DE CERDO", "11482": "CHOCOLATE SAHNE NUSS 3,750K (BARRA250G)", "11492": "VINAGRE DE VINO BLANCO", "11501": "PREMEZCLA DE MUFFINS", "11511": "POROTO GRANADO CONGELADO KG", "11512": "PASTA CHOCLO 1X16 KG", "11514": "HAMBURGUESA  STAR MEAT 113,5 GRS X42UN", "11521": "TE HIERBAS SURTIDAS 1X20UN", "11531": "COLCAFE  12 X 170 GRAMOS", "11532": "CAPUCCINO FRENCH VAINILLA 12 X 270 GRS", "11534": "GUATA CALLO", "11536": "BARRA CEREAL  15UN X CJ", "11537": "CEREAL CHOCAPIC 60UN X CJ", "11538": "LECHE INDIVIDUAL DE FRUTILLA 30UN X CJ", "11539": "LECHE BLANCA INDIVIDUAL", "11540": "LECHE INDIVIDUAL DE CHOCOLATE 30UN X CJ", "11541": "SURTIDO DE MARISCOS CONGELADO X 1 KILO", "11543": "CAFE COLCAFE CAPPUCCINO MOCCA 270GR X 12", "11551": "CHOCLO TROZO", "11552": "ESENCIA DE RON", "11561": "VACUNO TAPAPECHO V NAC.", "11571": "CHOCOLATE BELCOLADE BLANCO PURATOS", "11572": "CHOCOLATE BELCOLADE LECHE PURATOS", "11573": "CHOCOLATE BELCOLADE BITTER PURATOS", "11574": "CREMA MIX VEGETAL PURATOS", "11575": "PRALINE BELCOLADE ALMENDRA AVELLANA PURA", "11576": "FRUITFIL PINA PURATOS", "11577": "COCO RALLADO PURATOS", "11578": "COLDFIL NARANJA PURATOS", "11581": "GARBANZOS", "11582": "LIMON MEDIANO", "11583": "CARNE MOLIDA CORRIENTE 10%MG", "11584": "SALSA DE TOMATE PIZZA BIDON 120 KG", "11585": "SALSA ALFREDO", "11586": "TRUTO LARGO POLLO X 20 KG.", "11587": "PACK (MAYONESA, MOSTAZA Y KETCHUP)", "11588": "CIABATTA CJ X 10 KG", "11591": "PAPA CUBO VACIO 20X20 MM", "11593": "EMBUTIDO CHORIZO ESCUDERO", "11594": "MARRAQUETA", "11595": "POLLO FILETILLO GRILLADO", "11601": "BEBIDA LATA 350CC", "11602": "NECTAR ANDINA BOCA ANCHA", "11603": "COMPOTA DE FRUTAS", "11604": "PACK SERVICIO OSKU (1X200 UNIDADES)", "11605": "MEDIALUNA ARCO (1X200 UN) (NO UTILIZAR)", "11611": "MIX LECHUGAS (LOLOROSE,MARINA,ESPANOLA)", "11612": "CLARA LIQUIDA 1X5 KG BIDON", "11613": "PAN HOT DOG 15CM", "11614": "KETCHUP", "11615": "MAYONESA HELLMANNS", "11616": "MOSTAZA HELLMANNS", "11617": "MASA PIZZETA", "11621": "POLLO MOLIDO PREMIUN", "11631": "POLLO CUBO CONGELADO", "11641": "MEDIALUNA ICB 288 X 40GR FERMENTADA", "11652": "COCKTAIL DE FRUTA 6X3000 GR(NO UTILIZAR)", "11653": "ABASTERO CATEGORIA V", "11654": "LONGANIZA LO", "11661": "PECHUGA DESHUESADA CONGELADA", "11671": "POLLO TRUTO ALA NAC", "11672": "POLLO TRUTO LARGO NAC", "11673": "CABELLOS DE ANGEL", "11674": "SEMOLA", "11675": "JAMON PAVO COCIDO GRADO 1", "11677": "TOPPING DE FIAMBRE JAMON", "11681": "PAN BAGUETTE 110 GRS.  SESAMO", "11691": "DULCE DE MEMBRILLO (NO OCUPAR)", "11693": "REJILLA DULCE DE LECHE 80X47 GRS.", "11701": "ELABORADO BROCHETA MIXTA KG", "11702": "MIX ALITA COCIDA BBQ  10 KG", "11703": "MERMALADA SURTIDA VARIOS SABORES", "11704": "MORTADELA JAMONADA", "11706": "FRICA IDEAL 5  20 UNIDADES X BOLSA", "11708": "REJILLA DE FRAMBUESA 80X47 GRS.", "11709": "PAN MOLDE BANQUETE", "11710": "MERMALADA NARANJA TROZO X 5 KILOS", "11711": "DISCO MERENGUE CHICO 3 UNI. X BOLSA", "11712": "MASA A LA  PIEDRA FAMILIAR OREGANO", "11713": "MASA A LA  PIEDRA FAMILIAR INTEGRAL", "11722": "VACUNO HUESO CARNUDO", "11731": "MANZANA VERDES SLICE  1,9X6", "11732": "PAN FRICA 10 SIN SESAMO  45 GRAMOS", "11733": "PAN FRICA 11 SIN SESAMO  70 GRAMOS", "11734": "PAN FRICA 12 SIN SESAMO  90 GRAMOS", "11742": "CHURRASCO CONGELADO INTERFOLIADO", "11743": "QUESO CREMA FUNDIDO", "11744": "ALINO YOGURT ENELDO", "11746": "MARRASQUINOS ROJOS  GALON 6X3600/ 2,2 KG", "11747": "MARGARINA KG POTE 500 GR.", "11748": "EMPANADA DE PINO 220GR", "11751": "PERAS MITADES EN CONSERVA 6X3000 GR", "11753": "PROMOCION EMPANADA +BEBIDA", "11754": "CHORICILLO COCKTAIL", "11755": "TEGRAL SATIN CREAM CAKE 1X25 KG", "11756": "HARMONY FRIO 1X4,5 KG", "11757": "CREMIFIL MANGO 1X4,5 KG", "11758": "TEGRA CHOUX 1X15 KG", "11759": "FRUITFIL MANZANA 1X4 KG", "11760": "CREMFIL MANJAR 1X4,5 KG", "11761": "CREMFIL LIMON 1X4,5 KG", "11762": "CREMFIL LUCUMA 1X4,5 KG", "11763": "SURFIN (AZUCAR FLOR IMPALPABLE) 1X10 KG", "11771": "TEGRAL MOIST CHOCOLATE 1X20 KG", "11772": "FRUITFIL FRAMBUESA 1X4,5 KG", "11773": "GANACHE 1X4 KILOS", "11781": "JB MOST PPK FS 528X8 GRS", "11782": "JB KET FS 528X8 GRS", "11783": "HELLMANNS MAY FS 528X8 GRS", "11784": "PAN HOT DOG 19 CM (BOLSA 8 UNIDADES)", "11785": "PAN HOT DOG GRANEL", "11786": "PREMEZCLA PIZZA A LA PIEDRA 25 KG", "11791": "CEBOLLA POLVO", "11792": "VINAGRE BLANCO", "11801": "PULPA MANGO-PINA-NARANJA 1X10 KGIENDA", "11802": "PULPA MANGO-MARACUYA 1 X10 KGS TIENDA", "11803": "PULPA FRAMB-FRUTILLA-NARANJA 1X10 KGA", "11804": "PREMEZCLA CREAM CAKE (MUFFINS) ICB", "11805": "PREMEZCLA BROWNIE ICB", "11806": "LECHUGA LOLLO BIONDA HOJA 1KG", "11811": "PAN CIABATTA DE 12X5", "11813": "SACHET VINAGRETA OSKU  1X300", "11814": "MASA A LA PIEDRA FAMILIAR JUMBO", "11815": "CHORICILLO COCKTAIL PF", "11817": "TOMATE CUBITO - ORIGEN ITALIA", "11818": "PIMENTONES TIRA ROJA 6X1,8 KILOS DRENADO", "11819": "CARNE VACUNO COCIDA EN TIRAS (CAJA 20 KI", "11820": "ALBONDIGAS DE VACUNO COCIDA X KG (10X2KG", "11821": "CERDO PRIETAS NAC", "11831": "LOMO VETADO", "11851": "TRUFA CENCOSUD PURATOS 1 X 4 KILOS", "11852": "HAMBURGUESA VACUNO MR. BURGER", "11854": "AZUCAR DAMA BLANCA 1X50 KILOS (SACO)", "11855": "AZUCAR FLOR 1X25 KILOS (SACO)", "11861": "HUMITAS 1X1", "11862": "NUTELLA 15X 350 GRAMOS", "11871": "QUESO GOUDA LAMINADO 20G", "11881": "DESMOLDANTE", "11882": "PAN HALLULLAS 8 X10 CMS  / 85 GRAMOS", "11884": "FRAMBUESA CRUMBLE KILO", "11885": "YOGURT NATURAL 900 GRAMOS", "11886": "RICOTTA REPOSTERA BOLSA 5K", "11891": "PLANCHA HOJA 42X36X1 CM (1X3)", "11901": "BROCOLI PROCESADO", "11902": "HABAS CONGELADAS", "11903": "PORORO VERDE CONGELADO", "11904": "APIO MATA", "11905": "HAMBURGUESA VACUNO PREMIUM 100 GMS.", "11906": "PANCETA AHUMADA LAMINADA GRANEL", "11907": "SALSA BARBACUE", "11908": "QUESO  CHEDDAR LAMINADO", "11914": "DISCO HOJARASCA", "11921": "GALLETA VAINILLA", "11922": "COLORANTE VERDE MUSGO", "11931": "GALLETA COCO - SAN CAMILO", "11933": "GALLETA NARANJA - SAN CAMILO", "11934": "GALLETA MUSEO - SAN CAMILO", "11936": "DISCO HOJARASCA NRO 22", "11937": "BIZCOCHO  NEUTRO 1*25 KG", "11940": "CHOCLO DESGRANADO COCIDO", "11941": "CREMA PASTELERA NEUTRA FAST MIX 10X1,2 K", "11942": "CREMA VEGETAL MAURI ORO 12X1 KG.", "11951": "MIX 4 ESTACIONES", "11952": "TOMATE COCKTAIL", "11953": "ATUN LOMITO POUCH", "11954": "BOWL ENSAL003", "11961": "DECORGEL MARACUYA", "11962": "BAKEFIL GUINDA", "11963": "MERMELADA MARACUYA", "11964": "MERLUZA CONGELADA", "11965": "PAPAS NOISSETES (DUQUESA) 5X2,5KG", "11966": "COBERTURA CHOCOLATE MONEDAS", "11967": "FRUTA EN CONSERVA COCTEL KG"};
const PROVEEDORES = {"100000000": "Comercial Puerto Chico Ltda", "100000001": "Ingrid Herrera Alvarez", "100000002": "Soc. Com. S. Aguirre Z. y Cía Ltda", "100000003": "Cristian Naranjo y Compañia Ltda", "100000004": "Dist. y Com. Gary y Cia Ltda.", "100000005": "Renato Fuentes Jimenez", "100000006": "Confites Quijada y Compañia Ltda", "100000007": "Dialco Ltda", "100000008": "Augusto Serrano", "100000009": "Patricia Cruz Tapia", "100000010": "Enrique Rossel Barahona", "100000011": "Comercial Musri y Compañia Ltda", "100000012": "Ivan Arroyo Vera", "100000013": "Comercial Arroyito Ltda", "100000014": "Josaga EIRL", "100000015": "Transglobal Alimentos Spa", "100000016": "Distribuidora A&D SA", "100000017": "Distribuidora Garcia y Garces Ltda", "100000018": "Carlos A Millar Cid Dist. de Confit", "100000019": "Com. Arc-Dos Comercial E.I.R.L.", "100000020": "Com. y Dist. Amexpress Ltda.", "100000021": "Comercial CMC Sweet EIRL", "100000022": "ARTEL S.A.", "100000023": "Importadora Blobel S.A.", "100000024": "M Pin Chile SA", "100000025": "Empresas Carozzi S.A.", "100000026": "Comercial Costa SA", "100000027": "Caso y Cia S.A.I.C", "100000028": "Cartim SA", "100000029": "Cellstar Chile SA", "100000030": "CMPC Tissue SA", "100000031": "BAT Chile S.A.", "100000032": "Clan Ltda", "100000033": "Comercial Pinares Ltda", "100000034": "Ind. y Comercial Splum Ltda.", "100000035": "Comertex Representaciones SA", "100000036": "Comunicaciones Netglobalis SA", "100000037": "COPEC S.A.", "100000038": "Empresa periodistica La Cuarta SA", "100000039": "Empresa periodistica La Tercera SA", "100000040": "Empresa prensa asociada SA", "100000041": "Correos Chile SA", "100000042": "Cristian Bustos SA", "100000043": "Corpora Tresmontes SA", "100000044": "Diario El Sur SA", "100000045": "DIREC S.A.", "100000046": "Distribuidora Alfa SA", "100000047": "Dist.de Publicaciones CLC SA.", "100000048": "Distribuidora Molino SA", "100000049": "Embotelladora Chilena Unidas SA", "100000050": "Ediciones B Chile SA", "100000051": "Editores Unidos SA", "100000052": "Editorial Gestion Ltda.", "100000053": "Editorial Puelche SA", "100000054": "Editorial Televisa Chile SA", "100000055": "Empresa El Mercurio SAP", "100000056": "Empresa El Mercurio de Valparaiso S", "100000057": "Empresa Periodistica del Norte SA", "100000058": "Sociedad Comercial Enko Ltda", "100000059": "Evercrisp Snack Productos de Chile", "100000060": "Establecimiento de la Fuente SA", "100000061": "Gesnard Trading Company S.A", "100000062": "Importadora Cafe Do Brasil SAIC", "100000063": "Ideal SA", "100000064": "Distribuidora de Alimentos Vic Limi", "100000065": "Distribuidora y comercial C y M Ltd", "100000066": "Distribuidora Rodriguez y Reyes Ltd", "100000067": "Distribuidora y comercial Nueva", "100000068": "Sociedad distribuidora de alimentos", "100000069": "Distribuidora de alimentos Promar l", "100000070": "Distribuidora de alimentos Elpac lt", "100000071": "Distribuidora de alimentos Didmar l", "100000072": "Distribuidora de Alimentos Axe E.I.", "100000073": "Distribuidora de alimentos Maryser", "100000074": "Inversiones y Servicios Ilan ltda", "100000075": "Importadora y Distribuidora Jannis", "100000076": "Ramon Valente y Cia Ltda.", "100000077": "Comercializadora Kittyland Ltda", "100000078": "Jorge Rabie y Cia. S.A.", "100000079": "Inversiones la Mundial Ltda", "100000080": "Loreal Chile S.A.", "100000081": "Loteria de Concepcion", "100000082": "Comercial Canada Chemicals Ltda.", "100000083": "Maletas Chile Ltda", "100000084": "Mazapan Santa Ignacia Ltda", "100000085": "Mohana Hnos Ltda", "100000086": "Com. de Productos Nestle S.A.", "100000087": "Nestle Chile S.A.", "100000088": "Comercial e Industrial Novatec Ltda", "100000089": "Alimentos Nutrabien SA", "100000090": "Barra San Rafael", "100000091": "Nelson Yañez Yañez", "100000092": "Pibamour Sociedad Comercial y de", "100000093": "Polla Chilena de Beneficencia S.A.", "100000094": "Profac SA", "100000095": "PROMERCO S.A.", "100000096": "Salo SA", "100000097": "STONEX LTDA.", "100000098": "Velarde Hermanos SA", "100000099": "Victorinox Andes S.A.", "100000100": "Video Chile SA", "100000101": "Watts SA", "100000102": "Xerox de Chile SA", "100000103": "Agricola Santa Rita Ltda", "100000104": "Agrosuper Comercializadora de alime", "100000105": "Bacardi Martini Chile S.A", "100000106": "Aguas del Sur SA", "100000107": "Cervecera CCU Chile Ltda", "100000108": "Embotelladora Andina S.A", "100000109": "Embotelladora Iquique S.A", "100000110": "Coca Cola Embonor S.A.", "100000111": "Embotelladora Coca-Cola Polar S.A", "100000112": "Comercial Chacao SA", "100000113": "Compania Pisquera de Chile S.A", "100000114": "DISTRIBUCION Y EXCELENCIA S.A.", "100000115": "Compañia Comercializadora Nacional", "100000116": "Comercial Peumo Limitada", "100000117": "Hielo Fiesta SA", "100000118": "Claudia Bedregal Johnson", "100000119": "Distribuidora Troncoso Ltda", "100000120": "Distribuidora Frio Express Ltda", "100000121": "Importadora Euro Ltda", "100000122": "Comercial Santa Elena SA", "100000123": "Comercial Taboga S.A", "100000124": "Watt's Comercial S.A.", "100000125": "Sociedad Anonima Vina Santa Rita", "100000126": "Global Brands Chile S.A", "100000127": "Sadia Chile SPA", "100000128": "Cooperativa Agricola Pisquera Elqui", "100000129": "Minot Chile SA", "100000130": "Cia Agricola y Lechera Quillayes de", "100000131": "Vina San Pedro Tarapaca S.A.", "100000132": "Laboratorio Internacional Pharmacor", "100000133": "Alimentos y Frutos SA", "100000134": "Lagno, Leporati y Quer Ltda", "100000135": "Bredenmaster SpA", "100000136": "Soc. Com. Dist. Day-West Ltd.-", "100000137": "Imp. y Alim. ICB Food Service Ltda.", "100000138": "Oscar Eduardo Diaz muñoz", "100000139": "Juan Antonio Zambrano Bravo", "100000140": "Margarita del Carmen Perez Salazar", "100000141": "Sociedad Distribuidora El Rayo Ltda", "100000142": "Sociedad Comercial Muñoz e Hijos Lt", "100000143": "Emilio Silva Hijos y Compañia", "100000144": "Comercializadora Agrochequen Ltda", "100000145": "Dulcono Roma Limitada", "100000146": "Camilo Ferron Chile S.A", "100000147": "Comercial Ugarte y Cia Ltda", "100000148": "Puratos de Chile SPA", "100000149": "Sofruco Alimentos Limitada", "100000150": "Unilever Chile Limitada", "100000151": "Carnes Ñuble SA", "100000152": "CIAL Alimentos S.A.", "100000153": "Agrofoods Central Valley Chile S.A.", "100000154": "Aga S.A", "100000155": "Agroindustrias Diaguitas Limitada", "100000156": "Balmaceda Errazuriz Maria Magdalena", "100000157": "Caletas Marinas Comercializadora de", "100000158": "Coembal Chile S.A", "100000159": "Colorama S.A", "100000160": "Comercial Chau S.A.", "100000161": "Comercial Cobo Com Limitada", "100000162": "Comercial Coragro Limitada", "100000163": "COMERCIAL DEANDES LIMITADA", "100000164": "Comercializadora Ecuachile Limitada", "100000165": "Comercio Internacional Tagler S.A", "100000166": "Coop. agr. y lech. Union Ltda.", "100000167": "Diaz Latorre liberato y otros", "100000168": "Distribuidora Adelco Santiago Limit", "100000169": "Distribuidoras Aceitunas Ordoñez Li", "100000170": "Dole Chile SA", "100000171": "Ecolab SPA", "100000172": "Edapi S.A", "100000173": "Etiquetas Autoadhesivas Adhegraf Lt", "100000174": "Gestion Ecologica de Residuos S.A", "100000175": "International Paper Cartones S.A", "100000176": "Luisa Cristina Martinez Rivas", "100000177": "Multivac Chile S.A", "100000178": "New York Bagel S.A", "100000179": "Nutregg S.A", "100000180": "Pan D'or Chile S.A", "100000181": "Panificadora Las Rosas Chicas limit", "100000182": "Prodea S.A", "100000183": "Productos Fernandez S.A", "100000184": "Productos Plasticos H y C S.A.", "100000185": "Queensland Servicio y Mantenimiento", "100000186": "Salamanca Foods S.A", "100000187": "Scamusica Ltda", "100000188": "Sociedad Caprina Lonquimay limitada", "100000189": "Lacteos Mulpulmo", "100000190": "FOOD PACK SPA", "100000191": "Sociedad Comercial Ruiz y Mazry Lim", "100000192": "Trilogic S.A", "100000193": "Comercial CCU S.A.", "100000194": "Importadora Imprecin Ltda", "100000195": "Marinetti S.A", "100000196": "Ricardo Vargas", "100000197": "Fabrica Arco Alimentos", "100000198": "Biotec Chile SA", "100000199": "Barron Vieyra International SPA", "100000200": "Aba Image Service Ltda.", "100000201": "ABC Comercial Ltda", "100000202": "Abitek S.A.", "100000203": "Adhegrad Ltda.", "100000204": "Agencia de viajes Andina del Sur Lt", "100000205": "Alberto Mantelli y Cia.", "100000206": "Alexis Painemilla Picon", "100000207": "Alfacom", "100000208": "Goldstein y Cia Ltda", "100000210": "Andesmar Impresores S.A.", "100000211": "Araya y Oelckers Ltda.", "100000212": "Manuf. Elect. B Y P Ltda.", "100000213": "B.O. Foodservice S.A.", "100000214": "Bañados y Cia S.A.", "100000215": "Bash Seguridad S.A.", "100000216": "Brugnoli Arquitectos Asoc Ltda.", "100000217": "Compañia de Leasing Tattersal S.A.", "100000218": "SACI Falabella", "100000219": "Coasin Chile S.A.", "100000220": "Comercial Agrotecnica Ltda.", "100000221": "Comercial Gallardo y Wolff Ltda.", "100000222": "Comercial Grek-O-Lite Ltda.", "100000223": "Comercial Hernandez Morales Ltda.", "100000224": "Comercial Hispano Chilena", "100000225": "Comercial Napoleon", "100000226": "Comercial ZERO", "100000227": "Comtec", "100000228": "Constructora Jorge Carrasco", "100000229": "Claudio Vega Cid", "100000230": "Csl Electr. Y Comercial Systema Ltd", "100000231": "Telefonica Chile S.A.", "100000232": "DANIEL BARAHONA SILVA", "100000233": "Dantechnique", "100000234": "Dabed", "100000235": "Confecciones Sinetti Ltda", "100000236": "Distribuidora Perkins Chilena S.A.C", "100000237": "Distrimark Instore Media", "100000238": "Dorman", "100000239": "Ingenieria de Aceros Ltda", "100000240": "EASY S.A.", "100000241": "Comercial Serpan Ltda", "100000242": "MM", "100000243": "Editorial Pacifico S.A.", "100000244": "Emasa S.A.", "100000245": "Empresas Lipigas", "100000246": "Engatel S.A.", "100000247": "Eduardo H. Naduris Villalobos", "100000248": "Esina S.A.C.", "100000250": "Gloria Arellano Lopez", "100000251": "Ferreteria Amunategui S.A.", "100000252": "Gesa S.A.", "100000253": "Glasstech Ltda.", "100000254": "Globograf", "100000255": "Grupo Iluminacion Ltda.", "100000256": "Guest Gate Chile S.A.", "100000257": "Hanna Instruments Equipos Ltda.", "100000258": "Hasar Chile", "100000259": "Hertz Chile", "100000260": "Segundo Patricio Ibarra Muñoz", "100000261": "Imagen Color Ltda.", "100000262": "Imahe S.A.", "100000263": "Imega Ventus SpA", "100000264": "Impresion Uno", "100000265": "Impresora Digraf Ltda.", "100000266": "Inducrom S.A.", "100000267": "Ingenieria e Informatica Asoc Ltda.", "100000268": "Ingersoll Rand", "100000269": "Instalaciones Climaticas GCG", "100000270": "International Paper Chile S.A.", "100000271": "Jaime Riquelme Gonzalez", "100000272": "Jacuzzi Atilservice", "100000273": "Js Acri-neon Ltda.", "100000274": "Juan Ferrada F.", "100000275": "Juan Happ y Cia", "100000276": "Lanix Chile S.A.", "100000277": "Lechner y Cia. Ltda.", "100000278": "Linea Huno Ltda.", "100000279": "Macarena Calvo", "100000280": "Maria Ines Suarez de la Rivera", "100000281": "Mario Pizarro Martinez", "100000282": "Masmadera Limitada", "100000283": "Maximagen Ltda.", "100000284": "Mekano S.A.", "100000285": "Mellafe y Salas", "100000286": "Metrica", "100000287": "Fabricados Especiales Ltda.", "100000288": "Monica Vasquez", "100000289": "Moren & Moren Ltda.", "100000290": "Muebles Sur S.A.", "100000291": "MM", "100000292": "Musicwordl", "100000293": "Nancy Juacida Alcaino", "100000294": "NCR Chile Ltda.", "100000295": "NEC Chile S.A.", "100000296": "Netsecure", "100000297": "Nexcorp Industrial Ltda.", "100000298": "Nelson Pasten Fernandez", "100000299": "POP Diseño Ltda.", "100000300": "Papeles Imac Ltda.", "100000301": "Pelp International S.A.", "100000302": "Pesco S.A.", "100000303": "Peugeot Chile S.A.", "100000304": "Piedras El Gabino", "100000305": "EMILIO SOTO Y CIA. LTDA", "100000306": "PRECISION S.A", "100000307": "Precision Servicio S.A.", "100000308": "Prisa S.A.", "100000309": "Prodalum S.A.", "100000310": "Marketing y Publicidad Espinaca Ltd", "100000311": "Quattro Office International S.A.", "100000312": "ESTANTERIAS DEL MERCOSUR SPA", "100000313": "Grupo Reinax S.A", "100000314": "Rene Quezada San Martin", "100000315": "Ricardo Rodriguez y Cia Ltda", "100000316": "Roguemetal Ltda.", "100000317": "RENE OMAR HERNANDEZ VALDIVIA", "100000318": "Salva Austral S.A.", "100000319": "Saxamar S.A.", "100000321": "Servymaq", "100000322": "SET S.A.", "100000323": "Roberto Barraza Aldana", "100000324": "Sipack Soluciones Integrales", "100000325": "SKY Chile CPA", "100000326": "Serv. de Pintura y Techumbre Ltda.", "100000327": "Soporte y Proyectos en Computacion", "100000328": "Soluciones Prof.Integradas Ltda.", "100000329": "Steward S.A.", "100000330": "Triunfo Ltda.", "100000331": "Turismo Cocha S.A.", "100000332": "Ursus Trotter S.A.", "100000333": "Viveros Cinco Pinos Ltda.", "100000334": "Way Chile S.A", "100000335": "Juan Eduardo Ramos Ruiz", "100000336": "Marcos Caro y Cia Ltda.", "100000337": "Enrique Weis Lewin y Cia. Ltda.", "100000338": "Serv. Integ. Artic Ltda.", "100000339": "Comercial Tierra Nueva", "100000340": "Patricio Millacura", "100000341": "Velasco Leay y Cia Ltda.", "100000342": "Vidrios y Alum. Atenas Ltda.", "100000343": "Valitrans Ingeniería y Construcción", "100000344": "Mousse Plagass", "100000345": "Gloria Catalan Valin", "100000346": "Alexis Hernandez", "100000347": "Hernan Cordero Marin", "100000348": "Claudio Vergara", "100000349": "Besa Distribuidora", "100000350": "Automotriz Persico Hrnos. S.A.", "100000351": "Venthur Hnos y Cia Ltda", "100000352": "Enrique Solis Sanchez", "100000353": "Carmen Herrera", "100000354": "Mauricio Berrios", "100000355": "Julio Aviles", "100000356": "Aralum Construcciones", "100000357": "Heavenward Ascensores S.A", "100000358": "Patricio Rojas Mac Ginty y Cía. Ltd", "100000359": "Urbano Chile SPA", "100000360": "HDI Seguros S.A.", "100000361": "Servicios Equifax Chile Ltda.", "100000362": "Expresiones Creativas Gastronomicas", "100000363": "Alcopack SA", "100000364": "DPS CHILE COMERCIAL LIMITADA", "100000365": "Guillermo Del Carmen Segovia Cortez", "100000366": "Juan Carlos Venegas Carvajal.", "100000367": "Luis Hernan Vergara Gonzalez", "100000368": "Cima Group. S.A.", "100000369": "Rolando W.Fellmer Pfeiffer.", "100000370": "Chilectra S.A", "100000371": "Sergio Andres Arriagada Gonzalez.", "100000372": "Ingenieria y Construccion M y L. Lt", "100000373": "Empresa Elect de Antofagasta S.A", "100000374": "Beverage Air Interamericana Ltda.", "100000375": "Adm. Estaciones Servicio Serco Ltda", "100000376": "Maria Gricelda Almuna Flores.", "100000377": "Aguas Andina S.A.", "100000378": "GTD. Telesat. S.A.", "100000379": "Alfonso Fernando Cambiaso Varela.", "100000380": "Distrib. Nacional de Carnes Ltda.", "100000381": "Lorena Maribel Farias Astudillo", "100000382": "Transbank S.A.", "100000383": "Metrogas S.A.", "100000384": "Soc.Distr. y Com. Decam Pak Ltda.", "100000385": "Demetrio Enrique Silva Lopez.", "100000386": "ADELCO S.A. de Comercio.", "100000387": "AAAA Comercial Hispana Ltda.", "100000388": "Arco Impresores S.A.", "100000389": "Berguecio y Ebensperger Ltda", "100000390": "Telefonica Moviles Chile S.A.", "100000391": "Maria Ñanco Molina \" MAFA\"", "100000392": "Combustible Rapallo Ltda.", "100000393": "Telefonica del Sur.", "100000394": "Abastible S.A.", "100000395": "Enrique Adriazola Espinoza", "100000396": "Gas Sur S.A.", "100000397": "Sandro Alonzo Mendez Martinez", "100000398": "Jorge Marcelo Fierro Cabrera", "100000399": "Manuela D.Reuque Paillalef", "100000400": "Sitelcom E.I.R.L", "100000401": "Himce Ltda", "100000402": "Tecnoambiente S.A."};


/* Maestros vivos: parten con la muestra embebida y se reemplazan
   al cargar el Excel maestro completo (hojas CENTROS, SKU, PROVEEDORES).
   centros y skus almacenan objetos para soportar campos adicionales que
   se auto-completan en las planillas (orgVenta, unidadVenta, etc.). */
let MAESTROS = {
  centros: Object.fromEntries(Object.entries(CENTROS).map(([k, v]) => [k, { nombre: v, orgVenta: "BP01" }])),
  skus: Object.fromEntries(Object.entries(SKUS).map(([k, v]) => [k, { nombre: v, unidadVenta: "UN", unidadCondicion: "UN" }])),
  proveedores: { ...PROVEEDORES },
};

const norm = v => String(v ?? "").trim();
const code = v => { let s = norm(v); if (s.endsWith(".0")) s = s.slice(0, -2); return s; };
const fechaOK = v => /^\d{2}[-/]\d{2}[-/]\d{4}$/.test(norm(v));
/* Máscara DD-MM-AAAA: el usuario escribe solo dígitos y los guiones se insertan solos */
const maskFecha = v => {
  const d = String(v ?? "").replace(/\D/g, "").slice(0, 8);
  if (d.length <= 2) return d;
  if (d.length <= 4) return d.slice(0, 2) + "-" + d.slice(2);
  return d.slice(0, 2) + "-" + d.slice(2, 4) + "-" + d.slice(4);
};
const hoy = () => {
  const d = new Date();
  return String(d.getDate()).padStart(2, "0") + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + d.getFullYear();
};

/* ============================================================
   GUÍA DE SOLICITUDES (hoja MODIFICACIONES del archivo guía)
   Cada tipo de modificación indica qué planillas requiere.
   Las letras con portalId existen como mantenedor en este portal
   y se pueden preseleccionar; el resto se tramita por el canal
   habitual con su planilla correspondiente.
   ============================================================ */
const LETRAS = {
  "A":    { nombre: "Incorporación de Proveedor" },
  "B":    { nombre: "Incorporación Materiales SAP/SCKUBA" },
  "B.1":  { nombre: "Incorporación Materiales (Extras)" },
  "B.2":  { nombre: "Modificación Materiales SCKUBA" },
  "C":    { nombre: "Catalogación x Módulo", portalId: "catalogacion" },
  "C/C1": { nombre: "Catalogación (x Módulo o x Centro)", portalId: "catalogacion" },
  "C.1":  { nombre: "Catalogación x Centro", portalId: "catalogacion" },
  "C.2":  { nombre: "Vigencia Sckuba" },
  "D":    { nombre: "Registro Info (Inforecord)", portalId: "reginfo" },
  "E":    { nombre: "Precio Compra General o por Centro (PB00)", portalId: "preciocompra" },
  "F":    { nombre: "Cambio Unidad de Pedido" },
  "G":    { nombre: "Precio de Venta (VKP0)", portalId: "precio" },
  "G.1":  { nombre: "Precio de Venta Sckuba" },
  "H":    { nombre: "Comparación de Precios" },
  "I":    { nombre: "Borrado Libro de Pedido" },
  "J":    { nombre: "Libro de Pedido (Source List)", portalId: "libropedido" },
  "K":    { nombre: "Primer o Segundo Descuento (RA00 / RA01)", portalId: "descuento" },
  "L":    { nombre: "Flete (ZR01)", portalId: "flete" },
  "M":    { nombre: "Bloqueo", portalId: "bloqueo" },
  "N":    { nombre: "Desbloqueo", portalId: "bloqueo" },
  "Ñ":    { nombre: "Cambio Email de Pedido" },
  "O":    { nombre: "Emisor Factura" },
  "P":    { nombre: "Código Material Proveedor" },
  "Q":    { nombre: "Revalorización Material (MR21)" },
  "R":    { nombre: "Modifica Grupo de Artículo", portalId: "grupoarticulo" },
  "S":    { nombre: "Creación Combo" },
  "S2":   { nombre: "Combos Delivery" },
  "T":    { nombre: "Creación de Promoción" },
  "T2":   { nombre: "Promociones (complemento)" },
  "U":    { nombre: "Revalorización Set de Ventas" },
  "V":    { nombre: "Asignación de Impuesto (ZR02)", portalId: "impuesto" },
  "W":    { nombre: "Cambio de Descripción", portalId: "descripcion" },
  "X":    { nombre: "Nuevo EAN" },
  "X1":   { nombre: "Fuente Aprovisionamiento y Grupo Compra" },
};

const GUIA = [
  { n: 1,  nombre: "Incorporación de Proveedor", objetivo: "Nuevo proveedor para el sistema", letras: ["A"], nota: "Enviar ROL del Proveedor" },
  { n: 2,  nombre: "Incorporación de Material", objetivo: "Nuevo material para SAP", letras: ["B", "C/C1", "D", "E", "G", "J", "K", "L"], nota: "Enviar foto y medidas para Lay Out si corresponde" },
  { n: 3,  nombre: "Incorporación de Set de Ventas", objetivo: "Nuevo material con receta para la venta", letras: ["B", "C/C1", "G"] },
  { n: 4,  nombre: "Habilitación", objetivo: "Dejar activo un material ya creado en el sistema", letras: ["C/C1", "D", "E", "G", "J", "N"] },
  { n: 5,  nombre: "Cambio Precio Venta", objetivo: "Asignar nuevo precio de venta", letras: ["G"] },
  { n: 6,  nombre: "Cambio Precio Compra", objetivo: "Asignar nuevo precio de compra", letras: ["E"], nota: "Puede ser general o por centro" },
  { n: 7,  nombre: "Cambio de Proveedor", objetivo: "Asociar un material ya creado a un nuevo proveedor", letras: ["D", "E", "I", "J"], nota: "Puede ser general o por centro" },
  { n: 8,  nombre: "Bloqueo de Compra o Venta", objetivo: "Interrumpir la compra (Z1) o la venta (Z2) de un material", letras: ["M"] },
  { n: 9,  nombre: "Desbloqueo", objetivo: "Sacar bloqueo de compra o venta y habilitar nuevamente", letras: ["N", "G"] },
  { n: 10, nombre: "Carga de Descuento", objetivo: "Aplicar un descuento fijo (RA00) o temporal (RA01)", letras: ["K"] },
  { n: 11, nombre: "Carga de Flete", objetivo: "Aplicar un cargo por traslado al material", letras: ["L"] },
  { n: 12, nombre: "Cambio de Correo de Proveedor", objetivo: "Cambiar email de pedido de un proveedor", letras: ["Ñ"] },
  { n: 13, nombre: "Cambio Código Material Proveedor", objetivo: "Modificar códigos de proveedores en los materiales", letras: ["P"] },
  { n: 14, nombre: "Rectificar Emisor de Facturas", objetivo: "Modifica el proveedor que emitirá facturas de distribuidor", letras: ["O"] },
  { n: 15, nombre: "Cambio Unidad de Pedido", objetivo: "Modificación de unidad de embalaje y/o pedido", letras: ["F"] },
  { n: 16, nombre: "Cambio de Grupo de Artículo", objetivo: "Cambio de clasificación de material", letras: ["R"] },
  { n: 17, nombre: "Revalorización de Material", objetivo: "Modificar el costo de material", letras: ["Q"] },
  { n: 18, nombre: "Revalorización de Set de Venta", objetivo: "Correr proceso para que el set asuma el nuevo costo", letras: ["U"] },
  { n: 19, nombre: "Creación de Combo", objetivo: "Asociación de productos sin código de barras con un precio de venta determinado", letras: ["S", "S2"] },
  { n: 20, nombre: "Creación de Promociones", objetivo: "Asociar productos con código de barras a un precio de venta por un período determinado", letras: ["T", "T2"] },
  { n: 21, nombre: "Aplicar Impuesto", objetivo: "Asignar impuesto a material", letras: ["V"] },
  { n: 22, nombre: "Cambio de Descripción", objetivo: "Modifica la descripción del material", letras: ["W"] },
  { n: 23, nombre: "Nuevo EAN", objetivo: "Asociar nuevo código de barras a material ya creado", letras: ["X"] },
  { n: 24, nombre: "Habilitación Donación", objetivo: "Dejar activo material ya creado (Centro 6001)", letras: ["C/C1", "J", "N", "Q"] },
  { n: 25, nombre: "Cambio a Dueños de Inventario", objetivo: "Cambio de responsable de compras", letras: ["F", "D", "E", "C", "J", "X1"] },
];

const STATUS_VALIDOS = {
  "": "Desbloqueo",
  "Z1": "Bloqueo de compra",
  "Z2": "Bloqueo de compra y venta",
  "Z3": "Bloqueo centralizado",
  "Z4": "Bloqueo de Arcoalimentos",
};

/* ============================================================
   CLUSTERS DE LOCALES (por usuario)
   El usuario guarda listas de centros con nombre y puede escribir
   ese nombre en la columna Centro de cualquier planilla; el portal
   lo expande a sus centros al validar.
   ============================================================ */
let CLUSTERS = {}; // nombreLower -> { nombre, centros: [] }

const mapearClusters = obj =>
  Object.fromEntries(Object.entries(obj).map(([n, arr]) => [n.toLowerCase(), { nombre: n, centros: arr }]));

async function cargarClustersStorage() {
  const { data, error } = await supabase.from("clusters").select("nombre,centros");
  if (error || !data) return {};
  return Object.fromEntries(data.map(r => [r.nombre, r.centros]));
}

async function guardarClustersStorage(obj) {
  // Obtener clusters actuales del usuario para detectar eliminados
  const { data: existentes } = await supabase.from("clusters").select("id,nombre");
  const nombresNuevos = Object.keys(obj);
  const idsEliminar = (existentes ?? []).filter(r => !nombresNuevos.includes(r.nombre)).map(r => r.id);
  if (idsEliminar.length > 0) {
    await supabase.from("clusters").delete().in("id", idsEliminar);
  }
  const rows = Object.entries(obj).map(([nombre, centros]) => ({ nombre, centros }));
  if (rows.length > 0) {
    await supabase.from("clusters").upsert(rows, { onConflict: "user_id,nombre" });
  }
  return true;
}

/* Expande el contenido de la celda Centro: códigos numéricos pasan directo,
   tokens con texto se buscan como cluster. Deduplica dentro de la celda
   (un cluster puede traslaparse con códigos escritos a mano). */
function expandirCentros(texto) {
  const tokens = norm(texto).split(",").map(t => norm(t)).filter(Boolean);
  const out = [];
  tokens.forEach(t => {
    const c = code(t);
    if (/^\d+$/.test(c)) out.push(c);
    else {
      const cl = CLUSTERS[t.toLowerCase()];
      if (cl) cl.centros.forEach(x => out.push(code(x)));
      else out.push(t); // se reportará como cluster/centro inexistente
    }
  });
  return [...new Set(out)];
}

/* ============================================================
   LAS 4 PLANILLAS (estructura real del Sheet de Habilitación)

   claveUnica: campos que no pueden repetirse entre filas en este
   mantenedor. Si se omite, ninguna fila puede ser idéntica a otra
   (clave = todos los campos). Ejemplos de cómo se vería con reglas
   por mantenedor (PENDIENTE: confirmar reglas reales):
     claveUnica: ["centro", "sku"]                  → un registro por centro+SKU
     claveUnica: ["centro", "sku", "fechaInicio"]   → permite mismo centro+SKU en períodos distintos
   ============================================================ */
const PLANILLAS = [
  {
    id: "catalogacion",
    nombre: "Catalogación",
    desc: "Habilita uno o más SKU en uno o más centros.",
    icon: Boxes,
    userCols: [
      { key: "centro", label: "Centro", hint: "2003,2549 o cluster", w: 150, multi: true, centro: true },
      { key: "sku", label: "SKU", hint: "32", w: 100, sku: true },
      { key: "fechaInicio", label: "Fecha inicio", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "fechaFinal", label: "Fecha final", hint: "31-12-9999", w: 120, fecha: true, def: "31-12-9999" },
    ],
    autoCols: [
      { key: "nombreCentro", label: "Nombre centro", src: "centro_nombre" },
      { key: "descSku", label: "Descripcion SKU", src: "sku_desc" },
    ],
    salida: ["centro", "sku", "fechaInicio", "fechaFinal"],
    salidaLabels: ["Centro", "SKU", "Fecha inicio", "Fecha final"],
    ejemplo: { centro: "2003,2549", sku: "32", fechaInicio: "01-07-2026", fechaFinal: "31-12-9999" },
  },
  {
    id: "precio",
    nombre: "Precio Venta",
    desc: "Actualiza precio y condiciones de venta por centro.",
    icon: Tag,
    // Solo lo que el solicitante ingresa
    userCols: [
      { key: "centro", label: "Centro", hint: "2003,2549 o cluster", w: 150, multi: true, centro: true },
      { key: "sku", label: "SKU", hint: "32", w: 100, sku: true },
      { key: "fechaInicio", label: "Fecha inicio", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "fechaFinal", label: "Fecha final", hint: "31-12-9999", w: 120, fecha: true, def: "31-12-9999" },
      { key: "precio", label: "Precio", hint: "1290", w: 100, num: true },
      { key: "moneda", label: "Moneda", hint: "CLP", w: 90, def: "CLP" },
    ],
    // Auto-resueltos desde tablas maestras (visibles para ambos, solo lectura)
    autoCols: [
      { key: "nombreCentro",    label: "Nombre centro",          src: "centro_nombre" },
      { key: "descSku",         label: "Descripcion SKU",        src: "sku_desc" },
      { key: "unidadVenta",     label: "Unidad de venta",        src: "sku_unidadVenta" },
      { key: "unidadCondicion", label: "Unidad de condicion",    src: "sku_unidadCondicion" },
      { key: "orgVenta",        label: "Organizacion de venta",  src: "centro_orgVenta" },
    ],
    // Valores fijos internos: no visibles para el solicitante, van en el Excel SAP
    agentCols: [
      { key: "canal",          label: "Canal de distribucion", def: "02" },
      { key: "claseCondicion", label: "Clase de condicion",    def: "ZPVP" },
    ],
    salida: ["centro", "nombreCentro", "sku", "descSku", "fechaInicio", "fechaFinal", "unidadVenta", "precio", "unidadCondicion", "claseCondicion", "orgVenta", "canal", "moneda"],
    salidaLabels: ["Centro", "Nombre centro", "SKU", "Descripcion SKU", "Fecha inicio", "Fecha final", "Unidad de venta", "Precio", "Unidad de condicion", "Clase de condicion", "Organizacion de venta", "Canal de distribucion", "Moneda"],
    ejemplo: { centro: "2003", sku: "32", fechaInicio: hoy(), fechaFinal: "31-12-9999", precio: "1290", moneda: "CLP" },
  },
  {
    id: "bloqueo",
    nombre: "Bloqueo / Desbloqueo",
    desc: "Bloquea (Z1–Z4) o desbloquea SKU en centros. Status vacío = desbloqueo.",
    icon: Lock,
    userCols: [
      { key: "centro", label: "Centro", hint: "2003,2549 o cluster", w: 150, multi: true, centro: true },
      { key: "sku", label: "SKU", hint: "32", w: 100, sku: true },
      { key: "fechaInicio", label: "Fecha inicio", hint: "DD-MM-AAAA (vacío si desbloqueo)", w: 120, fecha: true, opcional: true, def: hoy() },
      { key: "status", label: "Status", hint: "Z1-Z4 o vacío", w: 110, status: true, opcional: true },
    ],
    autoCols: [
      { key: "nombreCentro", label: "Nombre centro", src: "centro_nombre" },
      { key: "descSku", label: "Descripcion SKU", src: "sku_desc" },
      { key: "significado", label: "Significado", src: "status_signif" },
    ],
    salida: ["centro", "nombreCentro", "sku", "descSku", "fechaInicio", "status"],
    salidaLabels: ["Centro", "Nombre centro", "SKU", "Descripcion SKU", "Fecha inicio", "Status"],
    ejemplo: { centro: "2003", sku: "32", fechaInicio: hoy(), status: "Z2" },
  },
  {
    id: "libropedido",
    nombre: "Libro de Pedido",
    desc: "Asigna proveedor y fuente de aprovisionamiento por centro.",
    icon: BookOpen,
    userCols: [
      { key: "centro", label: "Centro", hint: "2003,2549 o cluster", w: 150, multi: true, centro: true },
      { key: "sku", label: "SKU", hint: "32", w: 100, sku: true },
      { key: "fechaInicio", label: "Fecha inicio", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "fechaFinal", label: "Fecha final", hint: "31-12-9999", w: 120, fecha: true, def: "31-12-9999" },
      { key: "codProveedor", label: "Codigo proveedor", hint: "100000000", w: 130, prov: true },
      { key: "orgCompras", label: "Organizacion de compras", hint: "BP01", w: 150 },
      { key: "fuenteFija", label: "Fuente de aprovisionamiento fija", hint: "X o vacío", w: 170, opcional: true },
    ],
    autoCols: [
      { key: "nombreCentro", label: "Nombre centro", src: "centro_nombre" },
      { key: "descSku", label: "Descripcion SKU", src: "sku_desc" },
      { key: "nombreProveedor", label: "Nombre proveedor", src: "prov_nombre" },
    ],
    salida: ["centro", "nombreCentro", "sku", "descSku", "fechaInicio", "fechaFinal", "codProveedor", "nombreProveedor", "orgCompras", "fuenteFija"],
    salidaLabels: ["Centro", "Nombre centro", "SKU", "Descripcion SKU", "Fecha inicio", "Fecha final", "Codigo proveedor", "Nombre proveedor", "Organizacion de compras", "Fuente de aprovisionamiento fija"],
    ejemplo: { centro: "2003", sku: "32", fechaInicio: "01-07-2026", fechaFinal: "31-12-9999", codProveedor: "100000000", orgCompras: "BP01", fuenteFija: "X" },
  },
  {
    id: "preciocompra",
    nombre: "Precio de Compra",
    desc: "Carga el precio de compra (PB00), general o por proveedor.",
    icon: ShoppingCart,
    userCols: [
      { key: "condicion", label: "Condicion", hint: "PB00", w: 90, def: "PB00", cond: ["PB00"] },
      { key: "sku", label: "Codigo Material SAP", hint: "145873", w: 150, sku: true },
      { key: "proveedor", label: "Proveedor", hint: "100001173", w: 130, prov: true, opcional: true },
      { key: "orgCompras", label: "Organización de Compras", hint: "BP01", w: 160, def: "BP01" },
      { key: "precio", label: "Precio", hint: "669000", w: 110, num: true },
      { key: "inicioValidez", label: "Inicio Validez", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "finValidez", label: "Fin Validez", hint: "31-12-9999", w: 120, fecha: true, def: "31-12-9999" },
      { key: "base", label: "Base", hint: "100", w: 80, def: "100" },
      { key: "unidadBase", label: "Unidades Base", hint: "UN", w: 100, def: "UN" },
      { key: "moneda", label: "Moneda", hint: "CLP", w: 90, def: "CLP" },
    ],
    autoCols: [
      { key: "descSku", label: "Descripcion Material", src: "sku_desc" },
      { key: "nombreProveedor", label: "Nombre proveedor", src: "prov_nombre" },
    ],
    salida: ["condicion", "sku", "descSku", "proveedor", "nombreProveedor", "orgCompras", "precio", "inicioValidez", "finValidez", "base", "unidadBase", "moneda"],
    salidaLabels: ["Condicion", "Codigo Material SAP", "Descripcion Material", "Proveedor", "Nombre proveedor", "Organización de Compras", "Precio", "Inicio Validez", "Fin Validez", "Base", "Unidades Base", "Moneda"],
    ejemplo: { condicion: "PB00", sku: "145873", proveedor: "100001173", orgCompras: "BP01", precio: "669000", inicioValidez: "01-07-2026", finValidez: "31-12-9999", base: "100", unidadBase: "UN", moneda: "CLP" },
  },
  {
    id: "reginfo",
    nombre: "Registro Info",
    desc: "Crea el registro info (inforecord) que asocia material y proveedor.",
    icon: FileText,
    userCols: [
      { key: "sku", label: "Codigo Material SAP", hint: "145873", w: 150, sku: true },
      { key: "proveedor", label: "Proveedor", hint: "100000039", w: 130, prov: true },
      { key: "centro", label: "Centro", hint: "vacío = general", w: 140, multi: true, centro: true, opcional: true },
      { key: "orgCompras", label: "ORG_COMPRAS", hint: "BP01", w: 120, def: "BP01" },
      { key: "grpCompras", label: "GRP_COMPRAS", hint: "TPC / SIN", w: 120 },
      { key: "plazoEntrega", label: "PLAZO_ENTREGA", hint: "7", w: 120, num: true },
      { key: "unidadCompra", label: "Unidad Compra", hint: "CJ", w: 110 },
      { key: "codMatProv", label: "Código Mat Proveedor", hint: "CAPI0079", w: 150 },
      { key: "unidadBase", label: "Unidad Med Base", hint: "UN", w: 120, def: "UN" },
      { key: "factor", label: "Factor_Conversión", hint: "4", w: 130, num: true },
      { key: "denominador", label: "DENOMINADOR", hint: "1", w: 120, def: "1", num: true },
      { key: "tipoImpuesto", label: "Tipo Impuesto", hint: "C1", w: 110, def: "C1" },
      { key: "precio", label: "Precio (x cantidad base)", hint: "40000", w: 150, num: true },
      { key: "moneda", label: "Moneda", hint: "CLP", w: 90, def: "CLP" },
      { key: "inicioValidez", label: "Inicio Validez", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "cantidadBase", label: "Cantidad Base", hint: "100", w: 110, def: "100", num: true },
    ],
    autoCols: [
      { key: "descSku", label: "Descripción", src: "sku_desc" },
      { key: "nombreProveedor", label: "Nombre proveedor", src: "prov_nombre" },
      { key: "nombreCentro", label: "Nombre centro", src: "centro_nombre" },
    ],
    salida: ["sku", "descSku", "proveedor", "nombreProveedor", "centro", "nombreCentro", "orgCompras", "grpCompras", "plazoEntrega", "unidadCompra", "codMatProv", "unidadBase", "factor", "denominador", "tipoImpuesto", "precio", "moneda", "inicioValidez", "cantidadBase"],
    salidaLabels: ["Codigo Material SAP", "Descripción", "Proveedor", "Nombre proveedor", "Centro", "Nombre centro", "ORG_COMPRAS", "GRP_COMPRAS", "PLAZO_ENTREGA", "Unidad Compra", "Código Mat Proveedor", "Unidad Med Base", "Factor_Conversión", "DENOMINADOR", "Tipo Impuesto", "Precio", "Moneda", "Inicio Validez", "Cantidad Base"],
    ejemplo: { sku: "145873", proveedor: "100000039", centro: "", orgCompras: "BP01", grpCompras: "TPC", plazoEntrega: "7", unidadCompra: "CJ", codMatProv: "CAPI0079", unidadBase: "UN", factor: "4", denominador: "1", tipoImpuesto: "C1", precio: "40000", moneda: "CLP", inicioValidez: "01-07-2026", cantidadBase: "100" },
  },
  {
    id: "grupoarticulo",
    nombre: "Cambio Grupo Artículo",
    desc: "Reasigna el grupo de artículo (jerarquía) de un material.",
    icon: Layers,
    userCols: [
      { key: "sku", label: "Codigo Material SAP", hint: "178646", w: 150, sku: true },
      { key: "nuevoGrupo", label: "Nuevo Grupo de Articulo", hint: "02008CB06", w: 170 },
    ],
    autoCols: [
      { key: "descSku", label: "Descripcion Material", src: "sku_desc" },
    ],
    salida: ["sku", "descSku", "nuevoGrupo"],
    salidaLabels: ["Codigo Material SAP", "Descripcion Material", "Nuevo Grupo de Articulo"],
    ejemplo: { sku: "178646", nuevoGrupo: "02008CB06" },
  },
  {
    id: "descripcion",
    nombre: "Cambio Descripción",
    desc: "Modifica la descripción comercial de un material.",
    icon: Pencil,
    userCols: [
      { key: "sku", label: "Codigo Material SAP", hint: "147324", w: 150, sku: true },
      { key: "nuevaDescripcion", label: "Nueva", hint: "Máx. 40 caracteres", w: 220, maxLen: 40 },
    ],
    autoCols: [
      { key: "descSku", label: "Descripción actual", src: "sku_desc" },
    ],
    salida: ["sku", "descSku", "nuevaDescripcion"],
    salidaLabels: ["Codigo Material SAP", "Anterior", "Nueva"],
    ejemplo: { sku: "147324", nuevaDescripcion: "Ramita Mani 42g" },
  },
  {
    id: "flete",
    nombre: "Flete",
    desc: "Carga el cargo por traslado (ZR01) de un material por proveedor.",
    icon: Truck,
    userCols: [
      { key: "condicion", label: "Condicion", hint: "ZR01", w: 90, def: "ZR01", cond: ["ZR01"] },
      { key: "proveedor", label: "Proveedor", hint: "100000116", w: 130, prov: true },
      { key: "sku", label: "Codigo Material SAP", hint: "131764", w: 150, sku: true },
      { key: "importe", label: "Importe ($)", hint: "1855", w: 110, num: true },
      { key: "inicioValidez", label: "Inicio Validez", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "finValidez", label: "Fecha a", hint: "31-12-9999", w: 120, fecha: true, def: "31-12-9999" },
      { key: "moneda", label: "Moneda", hint: "CLP", w: 90, def: "CLP" },
      { key: "unidadBase", label: "Unidades Base", hint: "UN", w: 100, def: "UN" },
      { key: "base", label: "Base", hint: "100", w: 80, def: "100" },
    ],
    autoCols: [
      { key: "descSku", label: "Descripcion Material", src: "sku_desc" },
      { key: "nombreProveedor", label: "Nombre proveedor", src: "prov_nombre" },
    ],
    salida: ["condicion", "proveedor", "nombreProveedor", "sku", "descSku", "importe", "inicioValidez", "finValidez", "moneda", "unidadBase", "base"],
    salidaLabels: ["Condicion", "Proveedor", "Nombre proveedor", "Codigo Material SAP", "Descripcion Material", "Importe ($)", "Inicio Validez", "Fecha a", "Moneda", "Unidades Base", "Base"],
    ejemplo: { condicion: "ZR01", proveedor: "100000116", sku: "131764", importe: "1855", inicioValidez: "01-07-2026", finValidez: "31-12-9999", moneda: "CLP", unidadBase: "UN", base: "100" },
  },
  {
    id: "impuesto",
    nombre: "Impuesto Adicional",
    desc: "Asigna impuesto específico y/o adicional (ZR02) a un material.",
    icon: Percent,
    userCols: [
      { key: "condicion", label: "Condicion", hint: "ZR02", w: 90, def: "ZR02", cond: ["ZR02"] },
      { key: "sku", label: "Codigo Material SAP", hint: "131774", w: 150, sku: true },
      { key: "impuesto", label: "Impuesto (%)", hint: "18", w: 110, num: true },
      { key: "inicioValidez", label: "Inicio Validez", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "finValidez", label: "Fin Validez", hint: "31-12-9999", w: 120, fecha: true, def: "31-12-9999" },
    ],
    autoCols: [
      { key: "descSku", label: "Descripcion Material", src: "sku_desc" },
    ],
    salida: ["condicion", "sku", "descSku", "impuesto", "inicioValidez", "finValidez"],
    salidaLabels: ["Condicion", "Codigo Material SAP", "Descripcion Material", "Impuesto (%)", "Inicio Validez", "Fin Validez"],
    ejemplo: { condicion: "ZR02", sku: "131774", impuesto: "18", inicioValidez: "01-07-2026", finValidez: "31-12-9999" },
  },
  {
    id: "descuento",
    nombre: "Descuento",
    desc: "Carga descuento permanente (RA00) o temporal (RA01) por proveedor.",
    icon: TicketPercent,
    userCols: [
      { key: "condicion", label: "Condicion", hint: "RA00 / RA01", w: 110, cond: ["RA00", "RA01"] },
      { key: "proveedor", label: "Proveedor (pedido)", hint: "100000108", w: 140, prov: true },
      { key: "sku", label: "Codigo Material SAP", hint: "131774", w: 150, sku: true },
      { key: "importe", label: "Importe (Sin %)", hint: "15.33", w: 120, num: true },
      { key: "inicioValidez", label: "Inicio Validez", hint: "DD-MM-AAAA", w: 120, fecha: true, def: hoy() },
      { key: "finValidez", label: "Fecha a", hint: "31-12-9999", w: 120, fecha: true, def: "31-12-9999" },
    ],
    autoCols: [
      { key: "descSku", label: "Descripcion Material", src: "sku_desc" },
      { key: "nombreProveedor", label: "Nombre proveedor", src: "prov_nombre" },
    ],
    salida: ["condicion", "proveedor", "nombreProveedor", "sku", "descSku", "importe", "inicioValidez", "finValidez"],
    salidaLabels: ["Condicion", "Proveedor (pedido)", "Nombre proveedor", "Codigo Material SAP", "Descripcion Material", "Importe (Sin %)", "Inicio Validez", "Fecha a"],
    ejemplo: { condicion: "RA00", proveedor: "100000108", sku: "131774", importe: "15.33", inicioValidez: "01-07-2026", finValidez: "31-12-9999" },
  },
];

/* ============================================================
   VALIDACIÓN (se ejecuta solo al "Enviar a validación")
   ============================================================ */
const colKey = (planilla, flag) => planilla.userCols.find(c => c[flag])?.key;

function resolverAuto(planilla, src, row) {
  const get = flag => { const k = colKey(planilla, flag); return k ? row[k] : null; };
  switch (src) {
    case "centro_nombre":      return MAESTROS.centros[code(get("centro"))]?.nombre ?? null;
    case "centro_orgVenta":    return MAESTROS.centros[code(get("centro"))]?.orgVenta ?? null;
    case "sku_desc":           return MAESTROS.skus[code(get("sku"))]?.nombre ?? null;
    case "sku_unidadVenta":    return MAESTROS.skus[code(get("sku"))]?.unidadVenta ?? null;
    case "sku_unidadCondicion":return MAESTROS.skus[code(get("sku"))]?.unidadCondicion ?? null;
    case "prov_nombre":        return MAESTROS.proveedores[code(get("prov"))] ?? null;
    case "status_signif":      return STATUS_VALIDOS[norm(get("status")).toUpperCase()] ?? null;
    default: return null;
  }
}

function validarFila(planilla, row) {
  const e = [];
  planilla.userCols.forEach(c => {
    const v = norm(row[c.key]);
    if (c.centro) {
      if (!v) { e.push("Centro vacío"); return; }
      if (!MAESTROS.centros[code(v)]?.nombre) e.push(/^\d+$/.test(code(v))
        ? `Centro ${v} no existe en la base`
        : `Cluster o centro "${v}" no existe (revisa Gestor de locales)`);
      return;
    }
    if (c.sku) {
      if (!v) { e.push(`${c.label} vacío`); return; }
      if (!MAESTROS.skus[code(v)]?.nombre) e.push(`Material ${v} no está en el maestro`);
      return;
    }
    if (!v && !c.opcional) { e.push(`${c.label} vacío`); return; }
    if (!v) return;
    if (c.fecha && !fechaOK(v)) e.push(`${c.label}: use DD-MM-AAAA`);
    if (c.num) {
      const n = Number(v.replace(/[.\s$%]/g, ""));
      if (isNaN(n) || n <= 0) e.push(`${c.label} debe ser número mayor a 0`);
    }
    if (c.status && !["Z1", "Z2", "Z3", "Z4"].includes(v.toUpperCase())) e.push("Status inválido (use Z1-Z4, o vacío para desbloqueo)");
    if (c.prov && !MAESTROS.proveedores[code(v)]) e.push(`Proveedor ${v} no está en el maestro`);
    if (c.cond && !c.cond.map(x => x.toUpperCase()).includes(v.toUpperCase())) e.push(`${c.label} debe ser ${c.cond.join(" o ")}`);
    if (c.maxLen && v.length > c.maxLen) e.push(`${c.label} supera ${c.maxLen} caracteres`);
  });
  return e;
}

/* Toma las filas de la grilla, explota centros por coma, valida y autocompleta.
   Además detecta filas exactamente duplicadas (tras explotar centros):
   los mantenedores no aceptan dos registros idénticos.
   Devuelve { filas, porFila } donde porFila mapea índice de grilla -> errores agregados. */
function validarTodo(planilla, gridRows) {
  const filas = [];
  const porFila = {};
  const centroKey = colKey(planilla, "centro");
  gridRows.forEach((g, gi) => {
    if (planilla.userCols.every(c => !norm(g[c.key]))) return; // fila vacía: ignorar
    let lista;
    if (centroKey) { const cs = expandirCentros(g[centroKey]); lista = cs.length ? cs : [""]; }
    else lista = [null];
    lista.forEach(ce => {
      const row = centroKey ? { ...g, [centroKey]: ce } : g;
      const errores = validarFila(planilla, row);
      const auto = {};
      planilla.autoCols.forEach(c => { auto[c.key] = resolverAuto(planilla, c.src, row) ?? "—"; });
      const agent = {};
      (planilla.agentCols ?? []).forEach(c => {
        agent[c.key] = c.src ? (resolverAuto(planilla, c.src, row) ?? "—") : (c.def ?? "—");
      });
      filas.push({ gi, data: row, auto, agent, errores, ok: errores.length === 0 });
    });
  });

  // Unicidad: cada mantenedor define su clave en `claveUnica` (lista de keys).
  // Si no la define, la clave es la fila completa (todos los campos del usuario).
  const claves = planilla.claveUnica?.length
    ? planilla.userCols.filter(c => planilla.claveUnica.includes(c.key))
    : planilla.userCols;
  const firma = f => claves.map(c => code(f.data[c.key]).toUpperCase()).join("|");
  const etiquetaClave = claves.length === planilla.userCols.length
    ? "fila duplicada"
    : `combinación ${claves.map(c => c.label).join(" + ")} duplicada`;
  const vistos = new Map(); // firma -> primer índice en filas
  filas.forEach((f, i) => {
    const k = firma(f);
    if (vistos.has(k)) {
      const primera = filas[vistos.get(k)];
      f.errores.push(`${etiquetaClave.charAt(0).toUpperCase() + etiquetaClave.slice(1)} (igual a la fila ${primera.gi + 1} de la grilla)`);
      f.ok = false;
      if (!primera.errores.some(e => e.startsWith("Tiene duplicados"))) {
        primera.errores.push("Tiene duplicados más abajo: elimine las repeticiones");
        primera.ok = false;
      }
    } else {
      vistos.set(k, i);
    }
  });

  filas.forEach(f => {
    const errs = porFila[f.gi] ? new Set(porFila[f.gi]) : new Set();
    f.errores.forEach(x => errs.add(x));
    porFila[f.gi] = [...errs];
  });
  return { filas, porFila };
}

/* ============================================================
   EXCEL: plantilla, importación, exportación procesada
   ============================================================ */
function descargarPlantilla(planilla) {
  const headers = planilla.userCols.map(c => c.label);
  const ejemplo = planilla.userCols.map(c => planilla.ejemplo[c.key] ?? "");
  const ws = XLSX.utils.aoa_to_sheet([headers, ejemplo]);
  ws["!cols"] = headers.map(h => ({ wch: Math.max(h.length + 2, 16) }));
  const inst = XLSX.utils.aoa_to_sheet([
    ["Plantilla", planilla.nombre],
    [],
    ["1. No modifique la fila de encabezados."],
    ["2. La fila 2 es un ejemplo: reemplácela con sus datos."],
    ["3. En Centro puede escribir varios códigos separados por coma, o el nombre de uno de sus clusters del portal; se genera una fila por centro."],
    ["4. Nombre centro y Descripcion SKU los completa el portal al validar."],
    ["5. Si deja Fecha final vacía, el portal asume 31-12-9999 (sin vencimiento)."],
    planilla.id === "bloqueo" ? ["6. Status: Z1 compra, Z2 compra y venta, Z3 centralizado, Z4 Arcoalimentos, vacío = desbloqueo."] : [""],
  ]);
  inst["!cols"] = [{ wch: 85 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  XLSX.utils.book_append_sheet(wb, inst, "Instrucciones");
  XLSX.writeFile(wb, `Plantilla_${planilla.id}.xlsx`);
}

function importarExcel(planilla, workbook) {
  const sheet = workbook.Sheets["Datos"] || workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!raw.length) return { error: "El archivo está vacío." };
  const fileHeaders = raw[0].map(h => norm(h).toLowerCase());
  const expected = planilla.userCols.map(c => c.label.toLowerCase());
  const faltan = expected.filter(h => !fileHeaders.includes(h));
  if (faltan.length) return { error: `El archivo no corresponde a "${planilla.nombre}". Faltan columnas: ${faltan.join(", ")}.` };
  const idx = {};
  planilla.userCols.forEach(c => { idx[c.key] = fileHeaders.indexOf(c.label.toLowerCase()); });
  const rows = [];
  for (let i = 1; i < raw.length; i++) {
    const line = raw[i];
    if (line.every(v => norm(v) === "")) continue;
    const r = {};
    planilla.userCols.forEach(c => { r[c.key] = code(line[idx[c.key]]) && (c.key === "centro" || c.key === "sku" || c.prov) ? code(line[idx[c.key]]) : norm(line[idx[c.key]]) ; });
    // fechas: normalizar separadores y aplicar valor por defecto si viene vacía
    planilla.userCols.forEach(c => {
      if (!c.fecha) return;
      if (r[c.key]) r[c.key] = maskFecha(r[c.key]);
      else if (c.def) r[c.key] = c.def;
    });
    rows.push(r);
  }
  if (!rows.length) return { error: "No hay filas con datos bajo los encabezados." };
  return { rows };
}

function exportarSalida(planilla, filas) {
  const headers = planilla.salidaLabels;
  const aoa = [headers];
  filas.filter(f => f.ok).forEach(f => aoa.push(planilla.salida.map(k => f.data[k] ?? f.auto[k] ?? f.agent?.[k] ?? "")));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = headers.map(h => ({ wch: Math.max(h.length + 2, 14) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  XLSX.writeFile(wb, `${planilla.id}_validado_SAP.xlsx`);
}

const filaVacia = planilla => Object.fromEntries(planilla.userCols.map(c => [c.key, c.def ?? ""]));

/* ============================================================
   PERSISTENCIA DE SOLICITUDES
   Usa el almacenamiento del artifact (compartido: todos los
   usuarios del portal ven las solicitudes — solicitantes y
   equipo de Datos Maestros). Si no está disponible, cae a
   un registro en memoria de la sesión.
   ============================================================ */
const PREFIJO = "solicitudes:";

/* ---- Creación SKU ---- */
const TIPOS_SKU      = ["Retail", "Insumo", "Receta", "Combos", "Menaje", "Packaging"]; // lista completa (ref)
const TIPOS_SKU_BASE = ["Retail", "Insumo", "Menaje", "Packaging"]; // sin Receta ni Combos
const TIPOS_SIN_DIM  = new Set(["Receta", "Combos"]);
const TIPOS_SIN_FOTO = new Set(["Insumo", "Packaging"]);
const UNIDADES_DIM   = ["UN", "CJ", "DSP", "BL", "KG", "LT", "DOC", "PAR", "PAQ"];
const DIM_CAMPOS = [
  { key: "alto",       label: "Alto" },
  { key: "largo",      label: "Largo" },
  { key: "ancho",      label: "Ancho" },
  { key: "peso_neto",  label: "Peso neto" },
  { key: "peso_bruto", label: "Peso bruto" },
];
const emptyDim          = () => ({ unidad: "", alto: "", largo: "", ancho: "", peso_neto: "", peso_bruto: "" });
const emptyProductoCombo = () => ({ _id: Math.random().toString(36).slice(2), sku: "", delta_precio: "" });
const emptyPaso          = () => ({ _id: Math.random().toString(36).slice(2), nombre: "", obligatorio: "Si", pct_beneficio: "", productos: [emptyProductoCombo()] });
const emptySkuRow = () => ({
  _id: Math.random().toString(36).slice(2),
  id_creacion: "CSKU-" + new Date().toISOString().slice(0, 10).replace(/-/g, "") + "-" + Math.random().toString(36).slice(2, 6).toUpperCase(),
  nombre: "", descripcion: "", tipo: "",
  grupo_articulo: "", material_modelo: "",
  marca: "", fabricante: "",
  flag_planograma: false, flag_navegacion: false,
  factor_conversion: "",
  dim_primaria: emptyDim(), dim_secundaria: emptyDim(),
  insumos: [],
  pasos: [],
  fotos_planograma: [], fotos_navegacion: [],
  errores: [],
});
const memoria = []; // fallback

const ESTADOS = {
  "Enviada":    { color: "#5b8dee", bg: "rgba(91,141,238,0.1)" },
  "En proceso": { color: "#b25000", bg: "rgba(255,149,0,0.12)" },
  "Aprobada":   { color: "#248a3d", bg: "rgba(52,199,89,0.12)" },
  "Rechazada":  { color: "#c2271c", bg: "rgba(255,59,48,0.1)" },
  "Aplicada":   { color: "#515154", bg: "rgba(0,0,0,0.07)" },
};

async function guardarSolicitud(sol) {
  const { error } = await supabase.from("solicitudes").insert({
    folio:               sol.folio,
    solicitante_id:      sol.user_id,
    solicitante_nombre:  sol.solicitante,
    solicitante_email:   sol.email,
    estado:              sol.estado,
    planillas:           sol.planillas,
    historial:           sol.historial,
  });
  if (error) { console.error("guardar solicitud:", error); return "error"; }
  return "supabase";
}

async function listarSolicitudes(perfil) {
  let q = supabase.from("solicitudes").select("*").order("created_at", { ascending: false });
  if (perfil?.rol !== "datos_maestros") q = q.eq("solicitante_id", perfil?.id);
  const { data, error } = await q;
  if (error) { console.error("listar solicitudes:", error); return []; }
  return data || [];
}

async function cambiarEstado(folio, estado, motivo) {
  const { data: prev } = await supabase.from("solicitudes").select("historial").eq("folio", folio).single();
  const historial = [...(prev?.historial || []), { estado, fecha: new Date().toISOString() }];
  const upd = { estado, historial, updated_at: new Date().toISOString() };
  if (motivo) upd.motivo_rechazo = motivo;
  const { data, error } = await supabase.from("solicitudes").update(upd).eq("folio", folio).select().single();
  if (error) { console.error("cambiar estado:", error); return null; }
  return data;
}

const fmtFecha = iso => {
  try {
    return new Date(iso).toLocaleString("es-CL", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
  } catch { return iso; }
};

/* ============================================================
   COMPONENTE PRINCIPAL
   ============================================================ */
export default function PortalSAP() {
  const [session, setSession] = useState(undefined); // undefined = cargando
  const [perfil, setPerfil] = useState(null);
  const [vista, setVista] = useState("inicio"); // 'inicio' | 'nueva' | 'solicitudes' | 'clusters'

  const [clusters, setClusters] = useState({});
  const [seleccion, setSeleccion] = useState([]);
  const [authFlow, setAuthFlow] = useState(null); // null | 'set-password'

  // Auth: verificar sesión inicial y escuchar cambios
  useEffect(() => {
    const hash = window.location.hash;
    const esInvitacion = hash.includes("type=invite");

    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));

    const { data: { subscription } } = supabase.auth.onAuthStateChange((event, session) => {
      if (event === "PASSWORD_RECOVERY") {
        setAuthFlow("set-password");
        setSession(session);
      } else if (event === "SIGNED_IN" && esInvitacion && session) {
        setAuthFlow("set-password");
        setSession(session);
      } else {
        setSession(session);
      }
    });
    return () => subscription.unsubscribe();
  }, []);

  // Cargar perfil y maestros desde Supabase al iniciar sesión
  useEffect(() => {
    if (!session) return;
    (async () => {
      const [{ data: p }, { data: centros }, { data: skus }, { data: proveedores }] = await Promise.all([
        supabase.from("profiles").select("*").eq("id", session.user.id).single(),
        supabase.from("centros").select("codigo,nombre,org_venta"),
        supabase.from("skus").select("codigo,nombre,unidad_venta,unidad_condicion"),
        supabase.from("proveedores").select("codigo,nombre"),
      ]);
      if (p) setPerfil(p);
      if (centros && skus && proveedores) {
        MAESTROS = {
          centros: Object.fromEntries(centros.map(r => [r.codigo, { nombre: r.nombre, orgVenta: r.org_venta ?? "BP01" }])),
          skus: Object.fromEntries(skus.map(r => [r.codigo, { nombre: r.nombre, unidadVenta: r.unidad_venta ?? "UN", unidadCondicion: r.unidad_condicion ?? "UN" }])),
          proveedores: Object.fromEntries(proveedores.map(r => [r.codigo, r.nombre])),
        };
        setMaestrosInfo({ centros: centros.length, skus: skus.length, proveedores: proveedores.length });
      }
    })();
  }, [session]);

  React.useEffect(() => {
    (async () => {
      const obj = await cargarClustersStorage();
      CLUSTERS = mapearClusters(obj);
      setClusters(obj);
    })();
  }, []);

  const actualizarClusters = async obj => {
    CLUSTERS = mapearClusters(obj);
    setClusters(obj);
    await guardarClustersStorage(obj);
    // los clusters afectan la expansión de centros: invalidar validaciones previas
    setEstado(e => Object.fromEntries(Object.entries(e).map(([k, v]) =>
      [k, v.resultado ? { ...v, resultado: null, aviso: "Tus clusters cambiaron. Vuelve a enviar a validación." } : v])));
  };
  // estado[id] = { rows: [...], resultado: null | {filas, porFila}, aviso: null|string }
  const [estado, setEstado] = useState({});
  const [enviado, setEnviado] = useState(null);
  const [maestrosInfo, setMaestrosInfo] = useState(null);
  const [maestroOpen, setMaestroOpen] = useState(false);
  const fileInputs = useRef({});
  const maestroInput = useRef(null);

  const toggle = id => {
    setSeleccion(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
    setEstado(e => {
      if (e[id]) { const n = { ...e }; delete n[id]; return n; }
      const pl = PLANILLAS.find(p => p.id === id);
      return { ...e, [id]: { rows: [filaVacia(pl)], resultado: null, aviso: null } };
    });
  };

  const aplicarGuia = ids => {
    setSeleccion(prev => [...new Set([...prev, ...ids])]);
    setEstado(e => {
      const n = { ...e };
      ids.forEach(id => {
        if (!n[id]) {
          const pl = PLANILLAS.find(p => p.id === id);
          n[id] = { rows: [filaVacia(pl)], resultado: null, aviso: null };
        }
      });
      return n;
    });
    setTimeout(() => document.getElementById("paso-seleccion")?.scrollIntoView({ behavior: "smooth", block: "start" }), 80);
  };

  const setRows = (id, fn) => setEstado(e => ({
    ...e, [id]: { ...e[id], rows: fn(e[id].rows), resultado: null } // editar invalida la validación previa
  }));

  const editCell = (id, ri, key, val) => setRows(id, rows => rows.map((r, i) => i === ri ? { ...r, [key]: val } : r));
  const addRow = (id, pl) => setRows(id, rows => [...rows, filaVacia(pl)]);
  const dupRow = (id, ri) => setRows(id, rows => [...rows.slice(0, ri + 1), { ...rows[ri] }, ...rows.slice(ri + 1)]);
  const delRow = (id, ri) => setRows(id, rows => rows.length > 1 ? rows.filter((_, i) => i !== ri) : [filaVacia(PLANILLAS.find(p => p.id === id))]);

  const importar = useCallback((planilla, file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const res = importarExcel(planilla, wb);
        setEstado(e => {
          const cur = e[planilla.id];
          if (res.error) return { ...e, [planilla.id]: { ...cur, aviso: res.error } };
          const sinVacias = cur.rows.filter(r => planilla.userCols.some(c => norm(r[c.key])));
          return { ...e, [planilla.id]: { rows: [...sinVacias, ...res.rows], resultado: null, aviso: `${res.rows.length} filas importadas de ${file.name}. Presiona "Enviar a validación".` } };
        });
      } catch {
        setEstado(e => ({ ...e, [planilla.id]: { ...e[planilla.id], aviso: "No se pudo leer el archivo. ¿Es un .xlsx válido?" } }));
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const validar = (planilla) => {
    setEstado(e => {
      const cur = e[planilla.id];
      const res = validarTodo(planilla, cur.rows);
      const aviso = res.filas.length === 0 ? "Ingresa al menos una fila con datos antes de validar." : null;
      return { ...e, [planilla.id]: { ...cur, resultado: res.filas.length ? res : null, aviso } };
    });
  };

  const cargarMaestro = useCallback(file => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        // lee plano: codigo -> nombre (para proveedores)
        const lee = (hoja, kcol, vcol) => {
          const sh = wb.Sheets[hoja]; if (!sh) return null;
          const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });
          const m = {};
          for (let i = 1; i < rows.length; i++) {
            const k = code(rows[i][kcol]); const v = norm(rows[i][vcol]);
            if (k) m[k] = v;
          }
          return m;
        };
        // lee objeto: codigo -> { nombre, ...extraCols } con fallback a defaults
        const leeObj = (hoja, kcol, vcol, extras) => {
          const sh = wb.Sheets[hoja]; if (!sh) return null;
          const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });
          const m = {};
          for (let i = 1; i < rows.length; i++) {
            const k = code(rows[i][kcol]); const v = norm(rows[i][vcol]);
            if (!k) continue;
            const obj = { nombre: v };
            extras.forEach(({ key, col, def }) => { obj[key] = norm(rows[i][col]) || def; });
            m[k] = obj;
          }
          return m;
        };
        // CENTROS col 2 = org_venta | SKU col 4 = unidad_venta, col 5 = unidad_condicion
        const c = leeObj("CENTROS", 0, 1, [{ key: "orgVenta", col: 2, def: "BP01" }]);
        const s = leeObj("SKU", 0, 3, [{ key: "unidadVenta", col: 4, def: "UN" }, { key: "unidadCondicion", col: 5, def: "UN" }]);
        const p = lee("PROVEEDORES", 0, 1);
        if (!c || !s || !p) { setMaestrosInfo({ error: "El archivo no tiene las hojas CENTROS, SKU y PROVEEDORES." }); return; }
        MAESTROS = { centros: c, skus: s, proveedores: p };
        setMaestrosInfo({ centros: Object.keys(c).length, skus: Object.keys(s).length, proveedores: Object.keys(p).length });
        // invalidar validaciones previas: deben revalidar contra la base nueva
        setEstado(e => Object.fromEntries(Object.entries(e).map(([k, v]) => [k, { ...v, resultado: null, aviso: "Base maestra actualizada. Vuelve a enviar a validación." }])));
      } catch { setMaestrosInfo({ error: "No se pudo leer el archivo maestro." }); }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const plSel = PLANILLAS.filter(p => seleccion.includes(p.id));
  const totales = plSel.map(p => {
    const r = estado[p.id]?.resultado;
    return { id: p.id, ok: r ? r.filas.filter(f => f.ok).length : 0, err: r ? r.filas.filter(f => !f.ok).length : 0, validado: !!r };
  });
  const totalOk = totales.reduce((a, t) => a + t.ok, 0);
  const totalErr = totales.reduce((a, t) => a + t.err, 0);
  const todoValidado = plSel.length > 0 && totales.every(t => t.validado);
  const nombreSolicitante = norm(perfil?.nombre || "");
  const puedeEnviar = todoValidado && totalErr === 0 && totalOk > 0 && nombreSolicitante.length >= 3;

  const enviar = async () => {
    const folio = "SAP-" + new Date().getFullYear() + "-" + String(Math.floor(100000 + Math.random() * 899999));
    const sol = {
      folio,
      fecha: new Date().toISOString(),
      solicitante: nombreSolicitante,
      user_id: session?.user?.id,
      email: session?.user?.email,
      estado: "Enviada",
      historial: [{ estado: "Enviada", fecha: new Date().toISOString() }],
      planillas: plSel.map(p => ({
        id: p.id,
        nombre: p.nombre,
        filas: estado[p.id].resultado.filas.filter(f => f.ok).map(f => ({ data: f.data, auto: f.auto, ok: true })),
      })),
      totalFilas: totalOk,
    };
    const destino = await guardarSolicitud(sol);
    setEnviado({ folio, destino });
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const reiniciar = () => { setSeleccion([]); setEstado({}); setEnviado(null); };

  /* ---------- ÉXITO ---------- */
  if (enviado) {
    return (
      <AppShell vista={vista} setVista={v => { reiniciar(); setVista(v); }} perfil={perfil}>
        <main className="success-wrap">
          <div className="success-check"><CheckCircle2 size={64} strokeWidth={1.3} /></div>
          <h1 className="success-title">Solicitud enviada.</h1>
          <p className="success-sub">
            Folio <strong>{enviado.folio}</strong>. Quedó registrada en la sección Solicitudes, donde puedes seguir su estado.
            {enviado.destino === "memoria" && " (Almacenamiento persistente no disponible: el registro durará solo esta sesión.)"}
          </p>
          <div className="success-resumen">
            {plSel.map(p => (
              <div key={p.id} className="success-item">
                <p.icon size={18} /><span>{p.nombre}</span>
                <span className="success-count">{estado[p.id].resultado.filas.filter(f => f.ok).length} filas</span>
                <button className="link-export" onClick={() => exportarSalida(p, estado[p.id].resultado.filas)}><Download size={13} /> Excel validado</button>
              </div>
            ))}
          </div>
          <div className="success-botones">
            <button className="btn-soft" onClick={() => { reiniciar(); setVista("solicitudes"); }}><ClipboardList size={15} /> Ver solicitudes</button>
            <button className="btn-primary" onClick={reiniciar}><RotateCcw size={16} /> Nueva solicitud</button>
          </div>
        </main>
      </AppShell>
    );
  }

  /* ---------- AUTH GUARD ---------- */
  if (session === undefined) {
    return <div className="portal"><Estilos /><div className="login-loading">Cargando…</div></div>;
  }
  if (!session) {
    return <div className="portal"><Estilos /><VistaLogin /></div>;
  }
  if (authFlow === "set-password") {
    return <div className="portal"><Estilos /><VistaCrearContrasena onDone={() => setAuthFlow(null)} /></div>;
  }

  /* ---------- VISTAS SECUNDARIAS ---------- */
  if (vista === "inicio")      return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaInicio perfil={perfil} setVista={setVista} /></AppShell>;
  if (vista === "maestros")    return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaMaestros /></AppShell>;
  if (vista === "admin")       return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaAdmin /></AppShell>;
  if (vista === "planillas")   return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaMantenedorPlanillas /></AppShell>;
  if (vista === "nuevo-sku")    return <VistaCreacionSKU    perfil={perfil} session={session} vista={vista} setVista={setVista} />;
  if (vista === "nueva-receta") return <VistaCreacionReceta perfil={perfil} session={session} vista={vista} setVista={setVista} />;
  if (vista === "nueva-combo")  return <VistaCreacionCombo  perfil={perfil} session={session} vista={vista} setVista={setVista} />;
  if (vista === "mod-receta")   return <VistaModReceta      perfil={perfil} session={session} vista={vista} setVista={setVista} />;
  if (vista === "mod-combo")    return <VistaModCombo       perfil={perfil} session={session} vista={vista} setVista={setVista} />;
  if (vista === "solicitudes") return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaSolicitudes perfil={perfil} /></AppShell>;
  if (vista === "clusters")    return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaClusters clusters={clusters} onChange={actualizarClusters} /></AppShell>;
  if (vista === "ayuda")       return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaAyuda /></AppShell>;
  if (vista === "cvp")         return <AppShell vista={vista} setVista={setVista} perfil={perfil}><VistaCVP /></AppShell>;

  /* ---------- NUEVA SOLICITUD ---------- */
  const panelDerecho = (
    <PanelValidacion
      plSel={plSel} totales={totales} totalOk={totalOk} totalErr={totalErr}
      todoValidado={todoValidado} puedeEnviar={puedeEnviar}
      perfil={perfil} puedeEnviar={puedeEnviar}
      enviar={enviar} estado={estado}
    />
  );

  return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil} rightPanel={panelDerecho}>

      {/* ENCABEZADO */}
      <div className="content-header">
        <h1 className="content-title">Gestor de Solicitudes</h1>
        <p className="content-sub">y Ciclo de Vida del Producto</p>
      </div>

      {/* GUÍA DE SOLICITUDES */}
      <section className="seccion">
        <GuiaSolicitudes onAplicar={aplicarGuia} />
      </section>

      {/* PASO 1 */}
      <section className="seccion" id="paso-seleccion">
        <Paso n="1" t="Selecciona los cambios" s="Elige uno o varios tipos de solicitud, o usa la guía para que los marque por ti." />
        <div className="grid-tipos">
          {PLANILLAS.map(p => {
            const activo = seleccion.includes(p.id);
            return (
              <button key={p.id} className={"card-tipo" + (activo ? " activo" : "")} onClick={() => toggle(p.id)} aria-pressed={activo}>
                <div className="card-tipo-top">
                  <span className="card-icon"><p.icon size={22} strokeWidth={1.7} /></span>
                  <span className="card-check">{activo && <CheckCircle2 size={22} />}</span>
                </div>
                <h3>{p.nombre}</h3>
                <p>{p.desc}</p>
              </button>
            );
          })}
        </div>
      </section>

      {/* PASO 2 */}
      {plSel.length > 0 && (
        <section className="seccion fade-in">
          <Paso n="2" t="Ingresa los datos y envía a validar" s="Escribe en la grilla o importa el Excel. Los nombres y descripciones se completan al validar." />
          <div className="zona-cargas">
            {plSel.map(p => (
              <BloquePlanilla
                key={p.id} planilla={p} est={estado[p.id]}
                onEdit={(ri, k, v) => editCell(p.id, ri, k, v)}
                onAdd={() => addRow(p.id, p)}
                onDup={ri => dupRow(p.id, ri)}
                onDel={ri => delRow(p.id, ri)}
                onImport={f => importar(p, f)}
                onValidar={() => validar(p)}
                onExport={() => exportarSalida(p, estado[p.id].resultado.filas)}
                fileRef={el => (fileInputs.current[p.id] = el)}
                openFile={() => fileInputs.current[p.id]?.click()}
              />
            ))}
          </div>
        </section>
      )}

      <footer className="pie">Gestión de Datos Maestros · Las solicitudes validadas se procesan en el siguiente ciclo de carga en SAP.</footer>
    </AppShell>
  );
}

/* ============================================================
   BLOQUE POR PLANILLA: grilla editable + validación + resultados
   ============================================================ */
function BloquePlanilla({ planilla, est, onEdit, onAdd, onDup, onDel, onImport, onValidar, onExport, fileRef, openFile }) {
  const res = est?.resultado;
  const ok = res ? res.filas.filter(f => f.ok).length : 0;
  const err = res ? res.filas.length - ok : 0;
  const tieneDatos = est?.rows.some(r => planilla.userCols.some(c => norm(r[c.key])));

  return (
    <div className="bloque">
      <div className="bloque-head">
        <span className="card-icon sm"><planilla.icon size={18} strokeWidth={1.7} /></span>
        <div className="bloque-tit">
          <strong>{planilla.nombre}</strong>
          <span>
            {res
              ? <>Validado: <em className="t-ok">{ok} OK</em>{err > 0 && <> · <em className="t-err">{err} con error</em></>}</>
              : tieneDatos ? "Pendiente de validación" : "Sin datos aún"}
          </span>
        </div>
        <div className="bloque-acciones">
          <button className="btn-ghost" onClick={() => descargarPlantilla(planilla)}><Download size={14} /> Plantilla</button>
          <input type="file" accept=".xlsx,.xls" hidden ref={fileRef} onChange={e => { onImport(e.target.files[0]); e.target.value = ""; }} />
          <button className="btn-ghost" onClick={openFile}><Upload size={14} /> Importar Excel</button>
          <button className="btn-validar" onClick={onValidar} disabled={!tieneDatos}><ShieldCheck size={15} /> Enviar a validación</button>
        </div>
      </div>

      {est?.aviso && <div className="aviso-info"><AlertTriangle size={14} /> {est.aviso}</div>}

      {/* GRILLA DE INGRESO */}
      <div className="grilla-scroll">
        <table className="grilla">
          <thead>
            <tr>
              <th className="th-n">#</th>
              {planilla.userCols.map(c => <th key={c.key} style={{ minWidth: c.w }}>{c.label}</th>)}
              <th className="th-x"></th>
            </tr>
          </thead>
          <tbody>
            {est?.rows.map((r, ri) => {
              const errsFila = res?.porFila?.[ri];
              const filaConError = errsFila && errsFila.length > 0;
              const filaOK = res && errsFila && errsFila.length === 0;
              return (
                <React.Fragment key={ri}>
                  <tr className={filaConError ? "g-err" : filaOK ? "g-ok" : ""}>
                    <td className="td-n">
                      {filaConError ? <XCircle size={14} className="ic-err" /> : filaOK ? <CheckCircle2 size={14} className="ic-ok" /> : <Pencil size={12} className="ic-dim" />}
                    </td>
                    {planilla.userCols.map(c => (
                      <td key={c.key}>
                        <input
                          className="celda"
                          value={r[c.key]}
                          placeholder={c.hint}
                          onChange={e => onEdit(ri, c.key, c.fecha ? maskFecha(e.target.value) : e.target.value)}
                          inputMode={c.fecha || c.num ? "numeric" : undefined}
                          aria-label={`${c.label} fila ${ri + 1}`}
                        />
                      </td>
                    ))}
                    <td className="td-x">
                      <span className="fila-btns">
                        <button className="dz-clear dup" onClick={() => onDup(ri)} aria-label="Duplicar fila" title="Duplicar fila"><Copy size={13} /></button>
                        <button className="dz-clear" onClick={() => onDel(ri)} aria-label="Eliminar fila" title="Eliminar fila"><Trash2 size={13} /></button>
                      </span>
                    </td>
                  </tr>
                  {filaConError && (
                    <tr className="g-err-detalle">
                      <td></td>
                      <td colSpan={planilla.userCols.length + 1}>{errsFila.join(" · ")}</td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
      <button className="btn-addrow" onClick={onAdd}><Plus size={14} /> Agregar fila</button>

      {/* RESULTADO VALIDADO (vista expandida con autocompletados) */}
      {res && (
        <div className="tabla-wrap">
          <div className="tabla-head">
            <span>Resultado de validación: {res.filas.length} filas{res.filas.length > est.rows.filter(r => planilla.userCols.some(c => norm(r[c.key]))).length && " (centros explotados)"} · <strong className="t-ok">{ok} válidas</strong>{err > 0 && <> · <strong className="t-err">{err} con error</strong></>}</span>
            <button className="link-export" disabled={!ok} onClick={onExport}><Download size={13} /> Exportar Excel validado</button>
          </div>
          <div className="tabla-scroll">
            <table>
              <thead>
                <tr>
                  <th></th>
                  {planilla.userCols.map(c => <th key={c.key}>{c.label}</th>)}
                  {planilla.autoCols.map(c => <th key={c.key} className="th-auto">{c.label}</th>)}
                  <th>Observaciones</th>
                </tr>
              </thead>
              <tbody>
                {res.filas.map((f, i) => (
                  <tr key={i} className={f.ok ? "" : "row-err"}>
                    <td>{f.ok ? <CheckCircle2 size={15} className="ic-ok" /> : <XCircle size={15} className="ic-err" />}</td>
                    {planilla.userCols.map(c => <td key={c.key}>{norm(f.data[c.key]) || "—"}</td>)}
                    {planilla.autoCols.map(c => <td key={c.key} className="td-auto">{f.auto[c.key]}</td>)}
                    <td className="td-obs">{f.errores.join(" · ") || "OK"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function Paso({ n, t, s }) {
  return (
    <div className="paso-head">
      <span className="paso-num">{n}</span>
      <div><h2 className="paso-title">{t}</h2><p className="paso-sub">{s}</p></div>
    </div>
  );
}

/* ============================================================
   GUÍA DE SOLICITUDES: qué planillas enviar según el objetivo
   ============================================================ */
function GuiaSolicitudes({ onAplicar }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const [activa, setActiva] = useState(null); // n del tipo abierto

  const lista = GUIA.filter(g => {
    const t = (g.nombre + " " + g.objetivo).toLowerCase();
    return q.trim() === "" || q.toLowerCase().split(/\s+/).every(w => t.includes(w));
  });

  const detalle = GUIA.find(g => g.n === activa);
  const portalIds = detalle ? [...new Set(detalle.letras.map(l => LETRAS[l]?.portalId).filter(Boolean))] : [];
  const externas = detalle ? detalle.letras.filter(l => !LETRAS[l]?.portalId) : [];

  return (
    <div className={"guia-card" + (open ? " open" : "")}>
      <button className="maestro-head" onClick={() => setOpen(o => !o)}>
        <span className="maestro-ic guia-ic"><Sparkles size={18} /></span>
        <div className="maestro-txt">
          <strong>¿No sabes qué planillas enviar?</strong>
          <span>Busca lo que necesitas hacer y la guía te indica las planillas requeridas, marcando las disponibles en el portal.</span>
        </div>
        {open ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
      </button>

      {open && (
        <div className="guia-body">
          <input
            className="clu-input"
            value={q}
            placeholder="Busca: habilitación, precio, proveedor, combo, EAN…"
            onChange={e => { setQ(e.target.value); setActiva(null); }}
          />
          <div className="guia-lista">
            {lista.length === 0 && <div className="clu-vacio">Sin resultados para "{q}".</div>}
            {lista.map(g => (
              <button key={g.n} className={"guia-item" + (activa === g.n ? " on" : "")} onClick={() => setActiva(activa === g.n ? null : g.n)}>
                <span className="guia-num">{g.n}</span>
                <span className="guia-item-txt">
                  <strong>{g.nombre}</strong>
                  <em>{g.objetivo}</em>
                </span>
                <span className="guia-letras">{g.letras.join(" · ")}</span>
              </button>
            ))}
          </div>

          {detalle && (
            <div className="guia-detalle fade-in">
              <h3>{detalle.nombre}</h3>
              <p className="guia-obj">{detalle.objetivo}{detalle.nota && <span className="guia-nota"> — {detalle.nota}</span>}</p>
              <div className="guia-planillas">
                {detalle.letras.map(l => {
                  const info = LETRAS[l];
                  const enPortal = !!info?.portalId;
                  return (
                    <div key={l} className={"guia-row" + (enPortal ? " portal" : "")}>
                      <span className="guia-badge">{l}</span>
                      <span className="guia-row-nombre">{info?.nombre || l}</span>
                      {enPortal
                        ? <span className="guia-tag on"><CheckCircle2 size={12} /> En el portal</span>
                        : <span className="guia-tag">Canal habitual</span>}
                    </div>
                  );
                })}
              </div>
              {portalIds.length > 0 ? (
                <div className="guia-accion">
                  <button className="btn-validar" onClick={() => onAplicar(portalIds)}>
                    <Check size={15} /> Seleccionar {portalIds.length === 1 ? "la planilla disponible" : `las ${portalIds.length} planillas disponibles`}
                  </button>
                  {externas.length > 0 && (
                    <span className="guia-resto">Las planillas {externas.join(", ")} aún no están en el portal: envíalas por el canal habitual.</span>
                  )}
                </div>
              ) : (
                <div className="guia-resto solo">Esta solicitud aún no tiene mantenedores en el portal: envía {detalle.letras.length === 1 ? "la planilla" : "las planillas"} {detalle.letras.join(", ")} por el canal habitual.</div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   VISTA INICIO — landing/home
   ============================================================ */
function VistaInicio({ perfil, setVista }) {
  const MODULOS = [
    { id: "nueva",       label: "Nueva Solicitud",            desc: "Crea solicitudes de cambio en datos maestros SAP: precios, bloqueos, registros info y más.", icon: Plus,         color: "#5b8dee" },
    { id: "solicitudes", label: "Bandeja de Solicitudes",     desc: "Revisa el estado de tus solicitudes enviadas y descarga los archivos procesados.",             icon: Inbox,        color: "#34c759" },
    { id: "clusters",    label: "Gestor de Locales",          desc: "Organiza centros en clusters para agilizar la carga masiva de solicitudes.",                   icon: Boxes,        color: "#ff9500" },
    { id: "cvp",         label: "Ciclo de Vida del Producto", desc: "Procesos estandarizados de alta, modificación y baja de SKUs con matriz RACI y pasos.",        icon: Layers,       color: "#af52de" },
    { id: "ayuda",       label: "Centro de Ayuda",            desc: "Guías de solicitud, catálogo de planillas y ejemplos de cada módulo disponible.",              icon: Lightbulb,    color: "#ff6b35" },
  ];

  const nombre = perfil?.nombre?.split(" ")[0];

  return (
    <main className="sol-wrap inicio-wrap">
      {/* ── Hero oscuro split ── */}
      <div className="inicio-hero">
        {/* Lado izquierdo: visual abstracto */}
        <div className="inicio-hero-visual">
          <div className="inicio-planet-fire" />
          <div className="inicio-planet-earth" />
          <div className="inicio-stars" />
        </div>

        {/* Lado derecho: texto */}
        <div className="inicio-hero-text">
          <div className="inicio-logomark">N ·</div>
          <h1 className="inicio-title">
            Nexus se concibe como una iniciativa estratégica destinada a transformar y optimizar la gestión de datos maestros dentro de Arcoprime
          </h1>
          <p className="inicio-desc-blue">
            Su objetivo principal es establecer un marco de trabajo cohesivo y eficiente que garantice la integridad, precisión y accesibilidad de los datos esenciales para la operación y toma de decisiones en la compañía.
          </p>
          <p className="inicio-desc-orange">
            Este portal es la plataforma para la carga y el control de modificaciones de datos maestros.
          </p>
        </div>
      </div>

      <div className="inicio-logos">
        <img src="/logo-pronto.png" alt="Pronto"      className="inicio-logo" style={{ height: 84 }} />
        <img src="/logo-jv.png"     alt="Juan Valdez" className="inicio-logo" style={{ height: 110 }} />
        <img src="/logo-sb.svg"     alt="Sbarro"      className="inicio-logo" style={{ height: 64 }} />
        <img src="/logo-ap.png"     alt="Arcoprime"   className="inicio-logo" style={{ height: 64 }} />
      </div>

      {/* ── Módulos ── */}
      <h2 className="inicio-section-title">Acceso rápido</h2>
      <div className="inicio-modules">
        {MODULOS.map(m => {
          const Icon = m.icon;
          return (
            <button key={m.id} className="inicio-card" onClick={() => setVista(m.id)}>
              <div className="inicio-card-icon" style={{ background: m.color + "18", color: m.color }}>
                <Icon size={22} strokeWidth={1.8} />
              </div>
              <div className="inicio-card-body">
                <strong>{m.label}</strong>
                <span>{m.desc}</span>
              </div>
              <ChevronRight size={16} className="inicio-card-arrow" />
            </button>
          );
        })}
      </div>
    </main>
  );
}

/* ============================================================
   VISTA CENTRO DE AYUDA: flujos de procesos y catálogo de planillas
   ============================================================ */
function VistaAyuda() {
  const [expandido, setExpandido] = useState(null);

  return (
    <main className="sol-wrap">
      <div className="sol-head">
        <div>
          <h1 className="sol-title">Centro de Ayuda</h1>
          <p className="sol-sub">Entiende cómo fluyen las solicitudes y qué carga cada planilla.</p>
        </div>
      </div>

      {/* FLUJO DE PROCESAMIENTO */}
      <section className="seccion">
        <h2 className="ayuda-h2"><ArrowRight size={20} /> Flujo de una Solicitud</h2>
        <div className="flujo-pasos">
          <div className="paso paso-1">
            <div className="paso-num">1</div>
            <div className="paso-txt">
              <strong>Seleccionas planillas</strong>
              <span>Elige qué cambios necesitas (Catalogación, Precio, etc.) o déjate guiar por la guía de solicitudes.</span>
            </div>
          </div>
          <div className="paso-arrow"><ArrowDown size={18} /></div>
          <div className="paso paso-2">
            <div className="paso-num">2</div>
            <div className="paso-txt">
              <strong>Ingresas datos</strong>
              <span>Completa la grilla o importa un Excel. El portal te autocompleta nombres y descripciones de los maestros.</span>
            </div>
          </div>
          <div className="paso-arrow"><ArrowDown size={18} /></div>
          <div className="paso paso-3">
            <div className="paso-num">3</div>
            <div className="paso-txt">
              <strong>Validas</strong>
              <span>Haces clic en "Enviar a validación". El portal verifica centros, materiales, proveedores y detecta duplicados.</span>
            </div>
          </div>
          <div className="paso-arrow"><ArrowDown size={18} /></div>
          <div className="paso paso-4">
            <div className="paso-num">4</div>
            <div className="paso-txt">
              <strong>Envías</strong>
              <span>Ingresas tu nombre y presionas "Enviar solicitud". Se registra en el portal con un folio SAP-YYYY-NNNNNN.</span>
            </div>
          </div>
          <div className="paso-arrow"><ArrowDown size={18} /></div>
          <div className="paso paso-5">
            <div className="paso-num">5</div>
            <div className="paso-txt">
              <strong>Equipo Datos Maestros</strong>
              <span>Tu solicitud entra en estado "Enviada". El equipo la revisa, genera el archivo LSMW y la procesa en SAP.</span>
            </div>
          </div>
        </div>
      </section>

      {/* CATÁLOGO DE PLANILLAS */}
      <section className="seccion">
        <h2 className="ayuda-h2"><Boxes size={20} /> Catálogo de Planillas</h2>
        <p className="ayuda-intro">Cada planilla corresponde a un tipo de cambio en los datos maestros de SAP. Expande cualquiera para ver detalles, columnas y ejemplos.</p>
        
        <div className="ayuda-lista">
          {PLANILLAS.map(p => {
            const estaAbierta = expandido === p.id;
            const Icon = p.icon;
            return (
              <div key={p.id} className={"ayuda-item" + (estaAbierta ? " open" : "")}>
                <button className="ayuda-item-btn" onClick={() => setExpandido(estaAbierta ? null : p.id)}>
                  <span className="ayuda-item-ic"><Icon size={18} /></span>
                  <span className="ayuda-item-head">
                    <strong>{p.nombre}</strong>
                    <em>{p.desc}</em>
                  </span>
                  <span className="ayuda-item-toggle">{estaAbierta ? <ChevronUp size={18} /> : <ChevronDown size={18} />}</span>
                </button>
                {estaAbierta && (
                  <div className="ayuda-item-body fade-in">
                    <div className="ayuda-detalle">
                      <h4>¿Cuándo usarla?</h4>
                      <p>{p.desc}</p>
                    </div>
                    <div className="ayuda-detalle">
                      <h4>Columnas a completar</h4>
                      <div className="ayuda-cols">
                        {p.userCols.map(c => (
                          <div key={c.key} className="ayuda-col">
                            <span className="col-nombre">{c.label}</span>
                            <span className="col-tipo">
                              {c.centro && "Centro"}
                              {c.sku && "Material"}
                              {c.prov && "Proveedor"}
                              {c.fecha && "Fecha"}
                              {c.num && "Número"}
                              {c.status && "Status"}
                            </span>
                            <span className="col-hint">{c.hint}</span>
                            {c.def && <span className="col-def">Defecto: {c.def}</span>}
                            {c.opcional && <span className="col-opt">Opcional</span>}
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="ayuda-detalle">
                      <h4>Autocompletado</h4>
                      <div className="ayuda-auto">
                        {p.autoCols.length > 0 ? p.autoCols.map(c => (
                          <span key={c.key} className="auto-item"><Sparkles size={12} /> {c.label}</span>
                        )) : <span className="auto-item sin">Sin autocompletado</span>}
                      </div>
                    </div>
                    <div className="ayuda-detalle">
                      <h4>Ejemplo de uso</h4>
                      <div className="ayuda-ejemplo">
                        {Object.entries(p.ejemplo).map(([k, v]) => {
                          const col = p.userCols.find(c => c.key === k);
                          if (!col || v === "") return null;
                          return (
                            <div key={k} className="ej-fila">
                              <span className="ej-etiq">{col.label}:</span>
                              <span className="ej-val">{v}</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                    {p.salida && (
                      <div className="ayuda-detalle">
                        <h4>Columnas en el Excel de salida</h4>
                        <div className="ayuda-salida">
                          {p.salidaLabels.map((lbl, i) => (
                            <span key={i} className="salida-col">{lbl}</span>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </section>

      {/* VALIDACIONES GENERALES */}
      <section className="seccion">
        <h2 className="ayuda-h2"><ShieldCheck size={20} /> Validaciones que realiza el portal</h2>
        <div className="ayuda-validaciones">
          <div className="val-item">
            <h4><CheckCircle2 size={15} /> Centro / Cluster</h4>
            <p>Cada código o nombre de cluster debe existir en la base. Si escribes "Prontos RM" en la columna Centro, el portal busca ese cluster en Gestor de locales y lo expande a todos sus centros.</p>
          </div>
          <div className="val-item">
            <h4><CheckCircle2 size={15} /> Material (SKU)</h4>
            <p>El código debe estar en el maestro de SKU. El portal autocompleta la descripción si existe.</p>
          </div>
          <div className="val-item">
            <h4><CheckCircle2 size={15} /> Proveedor</h4>
            <p>Los códigos de proveedor se validan contra el maestro. El nombre se autocompleta.</p>
          </div>
          <div className="val-item">
            <h4><CheckCircle2 size={15} /> Fechas (DD-MM-AAAA)</h4>
            <p>Se valida el formato. El portal aplica máscara automática mientras escribes (01072026 → 01-07-2026). La fecha final por defecto es 31-12-9999 (vigencia indefinida).</p>
          </div>
          <div className="val-item">
            <h4><CheckCircle2 size={15} /> Números</h4>
            <p>Precios, importes, porcentajes deben ser mayores a 0. Se ignoran separadores de miles y símbolos.</p>
          </div>
          <div className="val-item">
            <h4><CheckCircle2 size={15} /> Duplicados</h4>
            <p>El portal detecta filas exactamente iguales (o duplicadas en la clave única de la planilla) tras expandir centros. Las rechaza y te indica cuál es la primera ocurrencia.</p>
          </div>
        </div>
      </section>

      {/* CONSEJOS */}
      <section className="seccion">
        <h2 className="ayuda-h2"><Lightbulb size={20} /> Consejos rápidos</h2>
        <ul className="ayuda-consejos">
          <li><strong>Carga de datos:</strong> puedes escribir manualmente o importar un Excel. Si importas, el portal ignora filas vacías.</li>
          <li><strong>Clusters de locales:</strong> ve a Gestor de locales, crea un cluster con los centros que repetirás, y luego escribe su nombre en la columna Centro. Se expande automáticamente.</li>
          <li><strong>Duplicación de filas:</strong> el botón "Duplicar fila" copia la que está debajo. Útil para cambios que afectan muchos centros o variantes.</li>
          <li><strong>Exportar validado:</strong> después de validar sin errores, haz clic en "Excel" para descargar el archivo procesado que va a SAP.</li>
          <li><strong>Solicitudes enviadas:</strong> ve a la pestaña "Solicitudes" para ver el estado, modificar estados si eres del equipo Datos Maestros, o descargar de nuevo el archivo.</li>
          <li><strong>Guía de solicitudes:</strong> si no sabes qué planillas enviar, usa la guía al principio: describe qué quieres lograr (habilitación, cambio de precio, etc.) y ella preselecciona las planillas del portal.</li>
        </ul>
      </section>
    </main>
  );
}


/* ============================================================
   DATOS CVP  — Sub-procesos v2.0
   ============================================================ */
const CVP_SUBPROCESOS = [
  /* ── Mapa To Be ─────────────────────────────────────────── */
  {
    id: "mapa", nombre: "Mapa To Be · Alta de SKU", color: "verde", version: "2.0", estado: "En Revisión",
    proposito: "Flujo principal del CVP: creación y habilitación completa de un nuevo SKU en SAP y SCKUBA, desde la evaluación comercial hasta la sincronización con todos los sistemas de tiendas.",
    alcance_inicio: "Identificación comercial de un producto nuevo.",
    alcance_fin: "SKU completamente habilitado en SAP, SCKUBA y tiendas.",
    contexto_preceden: "Ninguno (punto de entrada del CVP).",
    contexto_suceden: "Planogramación · Gestión y Modificación · Salida.",
    actores: [
      { rol: "Encargado Comercial (PM)", area: "Comercial", resp: "Evalúa viabilidad, define navegación, completa planilla de solicitud y genera ticket a DDMM." },
      { rol: "Analista Datos Maestros", area: "DDMM", resp: "Control de calidad de solicitud, creación de SKU en SAP, carga de condiciones y sincronización en SCKUBA." },
      { rol: "Focal de Abastecimiento", area: "Supply Chain", resp: "Coordinación de condiciones logísticas, flujo de abastecimiento y reposición automática." },
      { rol: "Focal de Planograma", area: "Planograma", resp: "Validación de planograma y cubicación del producto." },
      { rol: "Marketing", area: "Marketing", resp: "Provisión de imagen obligatoria cuando aplica navegación digital." },
      { rol: "Compras Corporativas", area: "Compras Corporativas", resp: "Lidera la validación de proveedores nuevos: audita documentos, aprueba y notifica." },
    ],
    raci: {
      cols: ["Comercial", "DDMM", "Supply Chain", "Planograma", "Marketing"],
      filas: [
        { act: "Evaluar viabilidad del producto", vals: ["R", "I", "C", "C", "C"] },
        { act: "Completar planilla técnica", vals: ["R", "I", "I", "I", "I"] },
        { act: "Creación básica SKU en SAP", vals: ["I", "R", "I", "I", "I"] },
        { act: "Carga de condiciones (Venta/Compra/Abast./Surtido)", vals: ["I", "R", "C", "I", "I"] },
        { act: "Sincronización SAP → SCKUBA", vals: ["I", "R", "I", "I", "I"] },
        { act: "Provisión de imagen (si aplica navegación)", vals: ["C", "I", "I", "I", "R"] },
        { act: "Validación de planograma", vals: ["A", "I", "I", "R", "I"] },
      ],
    },
    fases: [
      { nombre: "Evaluación comercial e ideación", pasos: [
        { n: 1, act: "Ideación de Producto", desc: "Se evalúa si es conveniente incorporar el producto analizando márgenes, costos y factibilidad logística.", resp: "Encargado Comercial" },
        { n: 2, act: "¿Producto nuevo?", desc: "Si ya existe → flujo de modificaciones o bajas. Si es nuevo → continúa.", resp: "Encargado Comercial", decision: true },
        { n: 3, act: "¿Existe proveedor?", desc: "Si ya está registrado → define alcance del SKU. Si no → gestionar registro del proveedor.", resp: "Encargado Comercial", decision: true },
      ]},
      { nombre: "Gestión de Proveedor (si aplica)", pasos: [
        { n: 4, act: "Solicitar creación de proveedor", desc: "Comercial recopila documentación comercial, tributaria y legal del nuevo proveedor.", resp: "Encargado Comercial" },
        { n: 5, act: "Enviar a validación", desc: "Envío formal del expediente a Compras Corporativas para auditoría y aprobación.", resp: "Encargado Comercial" },
        { n: 6, act: "Evaluación y Gestión de Proveedor", desc: "Compras Corporativas audita y aprueba (o rechaza con notificación).", resp: "Compras Corporativas", decision: true },
        { n: 7, act: "Notificar Encargado Comercial", desc: "Confirmación de que el código del proveedor está activo en los sistemas.", resp: "Compras Corporativas" },
      ]},
      { nombre: "Definición y envío de solicitud", pasos: [
        { n: 8, act: "¿Requiere solo código de SKU?", desc: "Parcial: para importaciones o compras anticipadas. Total: para venta directa en tienda.", resp: "Encargado Comercial", decision: true },
        { n: 9, act: "Creación Parcial (solo código SKU)", desc: "Datos mínimos (identidad, marca, jerarquía) para emitir código de barras de forma ágil.", resp: "Encargado Comercial" },
        { n: 10, act: "Gran Planilla (SKU Total)", desc: "Planilla completa: dimensiones, empaque, logística, precios, costos y flujos de transporte.", resp: "Encargado Comercial" },
        { n: 12, act: "Enviar solicitud a DDMM", desc: "Despacho formal del requerimiento (planilla técnica o ticket) hacia la mesa de Datos Maestros.", resp: "Encargado Comercial / Franquiciado" },
      ]},
      { nombre: "Procesamiento en SAP por DDMM", pasos: [
        { n: 13, act: "Recepción y control de calidad", desc: "Ingreso en Ticketera DDMM: revisión de consistencia, duplicados y campos vacíos.", resp: "DDMM" },
        { n: 14, act: "¿Información completa?", desc: "Si hay inconsistencias → notifica al solicitante para corregir. Si OK → carga de datos.", resp: "DDMM", decision: true },
        { n: 15, act: "Carga información básica SKU en SAP", desc: "Descripción técnica, clasificación y dimensiones de empaque.", resp: "DDMM" },
        { n: 16, act: "Cargar condiciones de venta", desc: "PVP, esquemas de impuestos y asignación por clústeres de tiendas.", resp: "DDMM" },
        { n: 17, act: "Cargar condiciones de compra", desc: "Registro Info en SAP: proveedor, costo neto unitario y lead times.", resp: "DDMM" },
        { n: 18, act: "Cargar condiciones de abastecimiento", desc: "Libro Pedido en SAP: flujo (Directo/Centralizado/Mixto) y parámetros de reposición automática.", resp: "DDMM" },
        { n: 19, act: "Cargar condiciones de surtido", desc: "Catalogación del SKU en la matriz de tiendas autorizadas y bloqueos por sucursal.", resp: "DDMM" },
      ]},
      { nombre: "Sincronización y cierre", pasos: [
        { n: 20, act: "Sincronización SAP → SCKUBA", desc: "El sistema propaga automáticamente la información a SCKUBA y terminales de tiendas.", resp: "DDMM" },
        { n: 21, act: "Informar a solicitante + célula de cierre", desc: "Cierre del ticket y alerta confirmando que el SKU está 100% activo y transaccional.", resp: "DDMM" },
        { n: 22, act: "¿Requiere planograma?", desc: "Si necesita ordenamiento en góndola → sub-proceso Planogramación. Si no → complementación final.", resp: "Encargado Comercial", decision: true },
        { n: 23, act: "Sub-proceso Planogramación", desc: "Análisis de cubicaje, dimensiones y layout. Entrega del Product List con ubicación física y digital.", resp: "Focal de Planograma" },
        { n: 24, act: "Completar configuración SKU", desc: "Carga final de atributos comerciales pendientes y reinyección al ciclo operativo.", resp: "Encargado Comercial" },
      ]},
    ],
  },

  /* ── Sub-proceso 1: Ideación ─────────────────────────────── */
  {
    id: "ideacion", numero: 1, nombre: "Ideación del Producto", color: "azul", version: "2.0", estado: "En Revisión",
    proposito: "Validar comercialmente la viabilidad de incorporar un producto nuevo al portafolio, asegurando coherencia entre indicadores de mercado, abastecimiento y estrategia comercial antes de solicitar su creación formal.",
    alcance_inicio: "Evaluación de indicadores comerciales por el Encargado Comercial.",
    alcance_fin: "Carga de datos básicos y condiciones del nuevo SKU por DDMM.",
    contexto_preceden: "Ninguno (punto de inicio del CVP).",
    contexto_suceden: "Alta de SKU · Gestión y Modificación.",
    actores: [
      { rol: "Encargado Comercial (PM)", area: "Comercial", resp: "Revisa indicadores, evalúa condiciones logísticas, solicita ficha al proveedor y define el tipo de SKU a crear." },
      { rol: "Analista Datos Maestros", area: "DDMM", resp: "Valida completitud del expediente del proveedor y ejecuta la carga en el maestro." },
      { rol: "Compras Corporativas", area: "Compras Corporativas", resp: "Registra y evalúa legal y financieramente al proveedor cuando es nuevo." },
      { rol: "I+D / Calidad", area: "I+D / Calidad", resp: "Gestiona auditoría técnica y revisa estándares sanitarios para productos Food." },
      { rol: "Marketing", area: "Marketing", resp: "Produce fotografía técnica en alta definición para navegación digital." },
    ],
    raci: {
      cols: ["Comercial PM", "Comercial Analista", "DDMM", "Compras Corp.", "I+D", "Calidad", "Marketing"],
      filas: [
        { act: "Revisar indicadores (comercial, abastecimiento, mercado)", vals: ["A", "R", "", "", "", "", ""] },
        { act: "Solicitar ficha de producto a proveedor", vals: ["A", "R", "", "", "", "", ""] },
        { act: "Gestión y Habilitación Corporativa (proveedor nuevo)", vals: ["", "", "C", "R", "", "", ""] },
        { act: "Validación de Consistencia DDMM", vals: ["", "", "R", "", "", "", ""] },
        { act: "Notificación de Registro Exitoso", vals: ["I", "", "I", "R", "", "", ""] },
        { act: "Auditoría Técnica y Calidad (Food)", vals: ["", "", "", "", "R", "R", ""] },
        { act: "Producción de Fotografía Digital", vals: ["", "", "", "", "", "", "R"] },
        { act: "Carga de Datos Básicos y Condiciones", vals: ["", "", "R", "", "", "", ""] },
      ],
    },
    fases: [
      { nombre: "Evaluación comercial", pasos: [
        { n: 1, act: "Revisar Indicadores Comerciales", desc: "Evaluación del desempeño comercial de la categoría y metas del portafolio.", resp: "Encargado Comercial" },
        { n: 2, act: "Revisar Indicadores Abastecimiento", desc: "Análisis de quiebres de stock y flujos de inventario actuales.", resp: "Encargado Comercial" },
        { n: 3, act: "Revisar Indicadores Mercado", desc: "Tendencias externas, comportamiento de competidores y oportunidades.", resp: "Encargado Comercial" },
        { n: 4, act: "Revisar Inclusiones con Proveedores", desc: "Posibilidad de incorporar el producto bajo acuerdos vigentes.", resp: "Encargado Comercial" },
        { n: 5, act: "Revisar condiciones Full", desc: "Viabilidad logística y almacenamiento bajo modalidad centralizada (Full).", resp: "Encargado Comercial" },
        { n: 6, act: "Revisar condiciones TCT", desc: "Flujo de distribución bajo el esquema de Transbordo Cruzado (TCT).", resp: "Encargado Comercial" },
        { n: 7, act: "Solicitar Ficha de producto a proveedor", desc: "Requerimiento de especificaciones técnicas directamente al proveedor.", resp: "Encargado Comercial" },
      ]},
      { nombre: "Gestión de proveedor", pasos: [
        { n: 8, act: "¿Existe proveedor?", desc: "Si SÍ → validación de categoría Food. Si NO → habilitación corporativa.", resp: "Encargado Comercial", decision: true },
        { n: 9, act: "Gestión y Habilitación Corporativa", desc: "Compras Corporativas registra y evalúa legal y financieramente al proveedor.", resp: "Compras Corporativas" },
        { n: 10, act: "Validación de Consistencia (DDMM)", desc: "DDMM revisa completitud del expediente y ejecuta carga en maestro de proveedores.", resp: "Datos Maestros" },
        { n: 11, act: "Notificación de Registro Exitoso", desc: "Tras emitirse el código en Espacio Proveedor, se notifica al área comercial.", resp: "Compras Corporativas / DDMM" },
      ]},
      { nombre: "Evaluación Food y navegación", pasos: [
        { n: 12, act: "¿Corresponde a Food?", desc: "Si SÍ → evalúa desarrollo técnico. Si NO → pasos de navegación digital.", resp: "Encargado Comercial", decision: true },
        { n: 13, act: "Auditoría Técnica y Calidad", desc: "I+D gestiona auditoría externa y Calidad evalúa estándares sanitarios.", resp: "I+D / Calidad" },
        { n: 14, act: "Revisión y Recepción del Desarrollo", desc: "I+D valida técnicamente y entrega documentación final al comercial.", resp: "I+D / Encargado Comercial" },
        { n: 15, act: "¿Aplica navegación?", desc: "Si SÍ → solicita fotografía. Si NO → definición de tipo de SKU.", resp: "Encargado Comercial", decision: true },
        { n: 16, act: "Producción de Fotografía Digital", desc: "Marketing produce la fotografía técnica en alta definición.", resp: "Marketing" },
        { n: 17, act: "Alineación Omnicanal", desc: "El comercial coordina la estrategia de navegación en plataforma digital.", resp: "Encargado Comercial" },
      ]},
      { nombre: "Creación del SKU", pasos: [
        { n: 18, act: "Definición de Tipo de SKU", desc: "SKU Parcial (planograma/importación directa) o SKU Total (venta directa).", resp: "Encargado Comercial", decision: true },
        { n: 19, act: "Carga de Datos en Maestro de Artículos", desc: "DDMM unifica la planilla técnica y ejecuta carga masiva de datos básicos y condiciones.", resp: "Datos Maestros" },
      ]},
    ],
  },

  /* ── Sub-proceso 2: Gestión y Modificación ──────────────── */
  {
    id: "gestion", numero: 2, nombre: "Gestión y Modificación de Producto", color: "naranja", version: "2.0", estado: "En Revisión",
    proposito: "Gestionar los cambios necesarios sobre productos activos del portafolio, asegurando que las modificaciones en condiciones comerciales, datos básicos, condiciones de abastecimiento y surtido sean ejecutadas correctamente en SAP.",
    alcance_inicio: "Identificación de una necesidad de modificación sobre un producto activo.",
    alcance_fin: "Cierre del proceso en sistema y notificación al proveedor.",
    contexto_preceden: "Alta de SKU · Ideación del Producto.",
    contexto_suceden: "Condiciones de Abastecimiento · Condiciones de Surtido · Salida de un Producto.",
    actores: [
      { rol: "Encargado Comercial (PM)", area: "Comercial", resp: "Identifica la modificación, completa la planilla correspondiente, envía la solicitud formal y gestiona escalamientos." },
      { rol: "Analista Datos Maestros", area: "DDMM", resp: "Recepciona la solicitud, evalúa el tipo de acción, aplica bloqueos SAP (Z1/Z2/Z3), carga condiciones, configura y cierra el proceso." },
      { rol: "Inventario / Finanzas", area: "Inventario / Finanzas", resp: "Recibe notificación de los cambios y confirma la notificación al proveedor." },
      { rol: "Focal de Abastecimiento", area: "Abastecimiento", resp: "Participa como Informado en las etapas clave del flujo." },
    ],
    raci: {
      cols: ["Comercial Solicitante", "PM Validador", "DDMM", "Inventario", "Abastecimiento"],
      filas: [
        { act: "Identificar necesidad de cambio", vals: ["R", "A", "I", "I", "I"] },
        { act: "Completar planilla de modificación", vals: ["R", "A", "C", "I", "I"] },
        { act: "Validar y enviar solicitud formal", vals: ["R", "A", "I", "I", "I"] },
        { act: "Evaluación técnica (Base/Actualización)", vals: ["C", "C", "R", "I", "I"] },
        { act: "Aplicación de Bloqueos SAP (Z1, Z2, Z3)", vals: ["A", "I", "R", "I", "I"] },
        { act: "Carga de datos de pedido y condiciones", vals: ["I", "I", "R", "I", "I"] },
        { act: "Configuración y catalogación en SAP", vals: ["I", "I", "R", "I", "I"] },
        { act: "Cierre de proceso en sistema", vals: ["I", "I", "R", "I", "I"] },
        { act: "Notificación al Proveedor", vals: ["I", "A", "I", "R", "I"] },
      ],
    },
    fases: [
      { nombre: "Solicitud de modificación", pasos: [
        { n: 1, act: "Identificar necesidad de modificación", desc: "El comercial detecta la necesidad: condiciones comerciales, datos básicos, abastecimiento o surtido.", resp: "Encargado Comercial" },
        { n: 2, act: "Solicitar modificación y definir tipo de cambio", desc: "Se selecciona el tipo de cambio y se bifurca el flujo hacia la planilla correspondiente.", resp: "Encargado Comercial" },
        { n: "3a", act: "Condiciones de venta → Precio venta", desc: "Completa planilla de cambio de precio venta.", resp: "Encargado Comercial" },
        { n: "3b", act: "Fuerza de compra → Precio compra", desc: "Completa planilla de cambio de precio compra.", resp: "Encargado Comercial" },
        { n: "3c", act: "Descuentos → Planilla descuento", desc: "Completa planilla de cambio de descuento.", resp: "Encargado Comercial" },
        { n: "3d", act: "Impuestos → Planilla impuesto", desc: "Completa planilla de cambio de impuesto.", resp: "Encargado Comercial" },
        { n: "3e", act: "Recargos → Planilla recargo", desc: "Completa planilla de cambio de recargo.", resp: "Encargado Comercial" },
        { n: "3f", act: "Datos Básicos → Dimensiones / Módulos / Fotografía", desc: "Actualiza información de dimensiones, módulos y/o fotografía del producto.", resp: "Encargado Comercial" },
        { n: "3g", act: "Condiciones de abastecimiento → Dato de pedido", desc: "Completa planilla del dato de pedido con sub-proceso correspondiente.", resp: "Encargado Comercial" },
        { n: "3h", act: "Condiciones de surtido → Sub-proceso surtido", desc: "Gestiona sub-proceso de Condiciones de Surtido.", resp: "Encargado Comercial" },
        { n: 4, act: "Enviar solicitud de modificación", desc: "Envío formal de la planilla completada hacia Datos Maestros para ejecución.", resp: "Encargado Comercial" },
      ]},
      { nombre: "Ejecución en SAP", pasos: [
        { n: 5, act: "Recepción y evaluación del tipo de acción", desc: "DDMM recepciona y determina si es bloqueo o habilitación.", resp: "Datos Maestros", decision: true },
        { n: "6a", act: "Bloquear productos (si aplica)", desc: "Bloqueo previo a realizar los cambios, según el tipo de modificación.", resp: "Datos Maestros" },
        { n: "6b", act: "Habilitar productos (si aplica)", desc: "Habilitación cuando el tipo de modificación lo requiere.", resp: "Datos Maestros" },
        { n: 7, act: "Seleccionar canal y aplicar bloqueos SAP", desc: "(a) Centralizado → Z1  (b) Portal Abastecimiento → Z2  (c) Tienda → Z3", resp: "Datos Maestros" },
        { n: 8, act: "Cargar datos de pedido en SAP", desc: "Carga de la información correspondiente al tipo de modificación.", resp: "Datos Maestros" },
        { n: 9, act: "Habilitar productos / Catalogación", desc: "Habilitación y/o catalogación en SAP según el tipo de cambio.", resp: "Datos Maestros" },
        { n: 10, act: "Cargar condiciones de venta / abastecimiento", desc: "Carga de condiciones actualizadas según tipo de modificación.", resp: "Datos Maestros" },
        { n: 11, act: "Configurar condiciones", desc: "Configuración de condiciones resultantes (incluye configuración de promociones si aplica).", resp: "Datos Maestros" },
        { n: 12, act: "Cargar condiciones de Abastecimiento finales", desc: "Carga de condiciones de abastecimiento finales.", resp: "Datos Maestros" },
        { n: 13, act: "Cerrar proceso en sistema", desc: "Cierre del proceso en el sistema (Cerrar Israel).", resp: "Datos Maestros" },
        { n: 14, act: "Informar cambios al Encargado Comercial", desc: "Comunicación de que los cambios han sido aplicados exitosamente.", resp: "Datos Maestros" },
        { n: 15, act: "Proveedor notificado (Inventario / Finanzas)", desc: "Inventario / Finanzas confirma que el proveedor fue notificado.", resp: "Inventario / Finanzas" },
      ]},
    ],
  },

  /* ── Sub-proceso 3: Salida ───────────────────────────────── */
  {
    id: "salida", numero: 3, nombre: "Salida del Producto", color: "rojo", version: "2.0", estado: "En Revisión",
    proposito: "Gestionar de forma coordinada y controlada la baja de un SKU del portafolio activo, asegurando que todos los sistemas, canales y áreas involucradas ejecuten las acciones de cierre de forma sincronizada.",
    alcance_inicio: "Identificación de un SKU local que debe salir del portafolio activo.",
    alcance_fin: "Envío de comunicación formal de salida y agotamiento de stock en tiendas.",
    contexto_preceden: "Gestión y Modificación de Producto.",
    contexto_suceden: "—",
    actores: [
      { rol: "Encargado Comercial", area: "Comercial", resp: "Lidera el proceso: identifica el SKU, evalúa la salida, coordina restricciones, gestiona stock remanente y envía comunicaciones corporativas." },
      { rol: "Marketing", area: "Marketing", resp: "Revisa vigencia de campañas visuales y autoriza/confirma eliminación del SKU en Menu Boards." },
      { rol: "Datos Maestros", area: "DDMM", resp: "Procesa la des-catalogación en el ERP central (Subproceso Condiciones de Surtido)." },
      { rol: "Abastecimiento", area: "Supply Chain", resp: "Anula parámetros logísticos automáticos y cierra órdenes de compra abiertas con proveedores." },
      { rol: "Omnicanalidad / e-Commerce", area: "Operaciones", resp: "Ejecuta la baja técnica del SKU en portales web, aplicaciones móviles y catálogos virtuales." },
      { rol: "Inventario / Total Loss", area: "Tiendas", resp: "Realiza seguimiento de stock remanente, ejecuta ajustes a cero y puede aprobar Bloqueo Total Z2 excepcional." },
      { rol: "Planograma", area: "Planograma", resp: "Actualiza los layouts físicos de tiendas, removiendo el producto de los planogramas de exhibición." },
    ],
    raci: {
      cols: ["Comercial", "Marketing", "Abast.", "e-Commerce", "Inv./TL", "Operac.", "DDMM", "Planograma"],
      filas: [
        { act: "Identificar SKU Local", vals: ["R", "I", "", "", "", "", "", ""] },
        { act: "¿Aplica salida de Menú Board?", vals: ["R", "A", "", "", "", "", "", "I"] },
        { act: "Notificar a Marketing", vals: ["R", "A", "", "I", "", "", "", ""] },
        { act: "Confirmar fecha eliminación menu board", vals: ["I", "R", "", "", "", "", "", ""] },
        { act: "Solicitar bloqueo de compra Z1", vals: ["R", "A", "I", "", "I", "", "", ""] },
        { act: "Cascada de Estrategias (Liquidación/Devolución/Donación)", vals: ["R", "A", "", "", "C", "", "", ""] },
        { act: "Generar Planilla de Bloqueo y enviar a DDMM", vals: ["R", "A", "", "", "I", "", "I", ""] },
        { act: "Subproceso Condiciones de Surtido", vals: ["I", "", "", "", "", "", "R", ""] },
        { act: "Informar Salida de Producto (Hito Corporativo)", vals: ["R", "A", "I", "I", "I", "I", "I", "I"] },
      ],
    },
    fases: [
      { nombre: "Fase I: Evaluación Inicial y Control de Menú Board", pasos: [
        { n: 1, act: "Identificar SKU Local", desc: "Identificación del código de producto local que debe iniciar el proceso de baja del portafolio.", resp: "Encargado Comercial" },
        { n: 2, act: "¿Aplica salida de Menú Board?", desc: "Si SÍ → notificar Marketing. Si NO → fase de evaluación de stock.", resp: "Encargado Comercial", decision: true },
        { n: 3, act: "Notificar a Marketing la salida de producto", desc: "Notificación formal a Marketing para medir el impacto promocional.", resp: "Encargado Comercial" },
        { n: 4, act: "Confirmar fecha de eliminación de menu board", desc: "Marketing revisa vigencia y confirma la fecha exacta de eliminación.", resp: "Marketing" },
        { n: 5, act: "Notificar a solicitante", desc: "Marketing formaliza y comunica la fecha límite de retiro al comercial.", resp: "Marketing" },
        { n: 6, act: "Recepcionar respuesta de marketing", desc: "El comercial procesa la confirmación de fechas y cierra el ramal promocional.", resp: "Encargado Comercial" },
      ]},
      { nombre: "Fase II: Evaluación de Inventario y Estrategias de Salida", pasos: [
        { n: 7, act: "¿Stock en 0?", desc: "Si NO → evaluar bloqueo de compra. Si SÍ → ir a Fase III.", resp: "Encargado Comercial", decision: true },
        { n: 8, act: "¿Tiene bloqueo de compra?", desc: "Si SÍ → evaluar estrategia de liquidación. Si NO → solicitar bloqueo Z1.", resp: "Encargado Comercial", decision: true },
        { n: 9, act: "Solicitar bloqueo de compra Z1", desc: "Restricción técnica Z1 en SAP para detener órdenes de compra futuras.", resp: "Encargado Comercial" },
        { n: 10, act: "Cascada de Estrategias (liquidación / devolución / donación)", desc: "El comercial evalúa secuencialmente las alternativas para evacuar el stock remanente.", resp: "Encargado Comercial" },
        { n: 11, act: "Ajustar stock a cero", desc: "Ajuste de inventario administrativo en el sistema maestro.", resp: "Encargado Comercial" },
        { n: 12, act: "Aplicar estrategia elegida", desc: "Salida física del stock remanente de las tiendas (regresa al paso 7).", resp: "Encargado Comercial" },
        { n: 13, act: "Ruta Excepcional de Stock", desc: "Si ninguna estrategia funciona → Bloqueo Total Z2 excepcional. DDMM revisa e Inventario aprueba.", resp: "Comercial / DDMM" },
      ]},
      { nombre: "Fase III: Clasificación Técnica y Gestión de Dependencias", pasos: [
        { n: 14, act: "Informar a áreas de inminente salida", desc: "Alerta técnica hacia equipos operativos para anticipar la desactivación total.", resp: "Encargado Comercial" },
        { n: 15, act: "¿SKU aplica Stock?", desc: "Si SÍ → evaluar combo. Si NO → tratamiento No Stock (combo/servicio).", resp: "Encargado Comercial", decision: true },
        { n: 16, act: "¿Es el único componente del combo?", desc: "Si SÍ → evaluar si bloquear el combo completo o Modificar Combo.", resp: "Encargado Comercial", decision: true },
        { n: 17, act: "¿Es componente Receta? / Modificación", desc: "Si SÍ → Revisar Receta para Modificar Receta o aplicar Bloqueo Total.", resp: "Encargado Comercial", decision: true },
        { n: 18, act: "Tratamiento de No Stock: ¿Es combo o servicio?", desc: "Si servicio en Ciclo Full/Tienda → esperar fin de ciclo. Si no → Bloqueo de Venta Z2.", resp: "Encargado Comercial" },
        { n: 19, act: "Tratamiento de Recetas Directas", desc: "Para recetas puras: validar surtido, exclusividad de insumos y Solicitar Eliminar en Navegación.", resp: "Encargado Comercial" },
      ]},
      { nombre: "Fase IV: Cierre Comercial y Ejecución Operativa en Paralelo", pasos: [
        { n: 20, act: "Generar Planilla de Bloqueo", desc: "El comercial consolida los bloqueos del SKU, combos y recetas en la planilla maestra oficial.", resp: "Encargado Comercial" },
        { n: 21, act: "Enviar solicitud a DDMM", desc: "Envío formal de planillas de modificación y bloqueo hacia DDMM.", resp: "Encargado Comercial" },
        { n: 22, act: "Subproceso Condiciones de Surtido", desc: "DDMM procesa la des-catalogación y restricciones técnicas en el maestro ERP.", resp: "Datos Maestros" },
        { n: 23, act: "Recepcionar Información DDMM", desc: "El comercial recibe confirmación técnica y ticket cerrado para auditoría.", resp: "Encargado Comercial" },
        { n: 24, act: "¿Ticket DDMM rechazado?", desc: "Si hay errores → solicitar corrección y volver a Fase I. Si OK → cierre definitivo.", resp: "Encargado Comercial", decision: true },
        { n: 25, act: "Informar Salida de Producto (Hito Corporativo)", desc: "Comunicación formal de baja definitiva del SKU a toda la organización.", resp: "Encargado Comercial" },
        { n: "26A", act: "Recepcionar información de salida — Abastecimiento", desc: "Cancelación definitiva de parámetros de reposición y órdenes logísticas abiertas.", resp: "Abastecimiento" },
        { n: "26B", act: "Recepcionar información de salida — Omnicanalidad", desc: "Baja técnica inmediata del SKU en canales de e-commerce y aplicaciones móviles.", resp: "Omnicanalidad / e-Commerce" },
        { n: "26C", act: "Recepcionar información — Inventario / Total Loss", desc: "Monitoreo de stock residual en tiendas y ajuste final a cero.", resp: "Inventario / Total Loss" },
        { n: "26D", act: "Recepcionar información y baja — Operaciones", desc: "Operaciones coordina con tiendas la salida definitiva del producto.", resp: "Operaciones" },
      ]},
    ],
  },

  /* ── Sub-proceso 4: Proveedor ────────────────────────────── */
  {
    id: "proveedor", numero: 4, nombre: "Evaluación y Creación de Proveedor", color: "morado", version: "2.0", estado: "En Revisión",
    proposito: "Habilitar formalmente a un nuevo proveedor dentro de los sistemas de la organización, asegurando que cumpla los requisitos comerciales, económicos y legales antes de ser aprobado.",
    alcance_inicio: "Recepción de solicitud de creación de proveedor por Compras Corporativas (originada desde Mi Viaje).",
    alcance_fin: "Notificación al solicitante del nuevo código de proveedor generado.",
    contexto_preceden: "Ideación de Producto.",
    contexto_suceden: "Alta de proveedor en SAP por parte de DDMM.",
    actores: [
      { rol: "Compras Corporativas", area: "Compras Corporativas", resp: "Lidera el proceso completo: recepciona, audita documentos, evalúa impedimentos, aprueba y notifica a DDMM y al solicitante." },
      { rol: "Jefe de Compras", area: "Compras Corporativas", resp: "Autoriza o rechaza la habilitación cuando se detecta un impedimento económico." },
      { rol: "Ética y Cumplimiento", area: "Ética y Cumplimiento", resp: "Recepciona y audita la solicitud ante impedimentos legales; aprueba o rechaza bajo criterios normativos." },
      { rol: "DDMM", area: "Datos Maestros", resp: "Recibe notificación de aprobación y procede con la alta del proveedor en SAP." },
      { rol: "Encargado Comercial", area: "Comercial", resp: "Origina la solicitud en Mi Viaje y es receptor final de la notificación de resultado." },
    ],
    raci: {
      cols: ["Enc. Comercial", "Compras Corp.", "Jefe Compras", "Ética/Cumpl.", "DDMM", "Abastecimiento"],
      filas: [
        { act: "Origen de la necesidad (Trigger en Mi Viaje)", vals: ["R", "I", "I", "I", "I", "I"] },
        { act: "Recepcionar y revisar formulario de solicitud", vals: ["I", "R", "A", "I", "I", "I"] },
        { act: "Revisar completitud de documentos legales", vals: ["I", "R", "A", "I", "I", "I"] },
        { act: "Evaluar impedimento económico", vals: ["I", "R", "A", "I", "I", "I"] },
        { act: "Revisar solicitud y autorizar (Jefe Compras)", vals: ["I", "I", "R", "I", "I", "I"] },
        { act: "Evaluar impedimento legal", vals: ["I", "R", "A", "I", "I", "I"] },
        { act: "Revisar y aprobar (Ética y Cumplimiento)", vals: ["I", "I", "I", "R", "I", "I"] },
        { act: "Aprobar Proveedor", vals: ["I", "R", "A", "C", "I", "I"] },
        { act: "Notificar a DDMM", vals: ["I", "R", "A", "I", "I", "I"] },
        { act: "Crear proveedor (alta en SAP)", vals: ["I", "I", "I", "I", "R", "I"] },
        { act: "Notificar al solicitante (cierre del proceso)", vals: ["I", "R", "A", "I", "I", "I"] },
      ],
    },
    fases: [
      { nombre: "Revisión de solicitud y documentos", pasos: [
        { n: 1, act: "Recepcionar solicitud de creación de proveedor", desc: "El flujo se detona al ingresar un requerimiento en Mi Viaje.", resp: "Compras Corporativas" },
        { n: 2, act: "Revisar completitud de formulario", desc: "Auditoría visual y validación de datos obligatorios en Mi Viaje.", resp: "Compras Corporativas" },
        { n: 3, act: "Revisar completitud de documentos legales", desc: "Verificación de documentación técnica, tributaria y legal adjunta.", resp: "Compras Corporativas" },
      ]},
      { nombre: "Evaluación de impedimentos", pasos: [
        { n: 4, act: "¿Existe Impedimento Económico?", desc: "Si NO → evaluar impedimentos legales. Si SÍ → enviar a Jefe de Compras.", resp: "Compras Corporativas", decision: true },
        { n: 5, act: "Enviar a Jefe de Compras", desc: "Transferencia formal del ticket por impedimento financiero detectado.", resp: "Compras Corporativas" },
        { n: 6, act: "Revisar solicitud (Jefe de Compras)", desc: "Análisis comercial y financiero de los antecedentes de riesgo.", resp: "Jefe de Compras" },
        { n: 7, act: "¿Autoriza la creación?", desc: "Si NO → rechazar y notificar al solicitante. Si SÍ → retorna al camino principal.", resp: "Jefe de Compras", decision: true },
        { n: 8, act: "¿Existe Impedimento Legal?", desc: "Si NO → aprobación final. Si SÍ → enviar a Ética y Cumplimiento.", resp: "Compras Corporativas", decision: true },
        { n: 9, act: "Enviar a Ética y Cumplimiento", desc: "Escalamiento formal hacia la auditoría normativa.", resp: "Compras Corporativas" },
        { n: 10, act: "Revisar solicitud (Ética y Cumplimiento)", desc: "Análisis exhaustivo del perfil del tercero bajo marcos legales.", resp: "Ética y Cumplimiento" },
        { n: 11, act: "¿Aprueba la solicitud?", desc: "Si NO → rechazar. Si SÍ → retorna al camino principal para aprobación.", resp: "Ética y Cumplimiento", decision: true },
      ]},
      { nombre: "Aprobación y alta en SAP", pasos: [
        { n: 12, act: "Aprobar Proveedor", desc: "Validación y adjudicación definitiva del tercero en Mi Viaje.", resp: "Compras Corporativas" },
        { n: 13, act: "Notificar a DDMM", desc: "Instrucción técnica formal hacia DDMM para aprovisionar en sistemas maestros.", resp: "Compras Corporativas" },
        { n: 14, act: "Crear proveedor (alta en SAP)", desc: "Alta en SAP: estructuras bancarias, de compras e impositivas. Generación de código único.", resp: "DDMM" },
        { n: 15, act: "Recepcionar Nuevo Código de Proveedor", desc: "Toma de conocimiento del identificador oficial emitido para futuras transacciones.", resp: "Compras Corporativas" },
        { n: 16, act: "Notificar al Solicitante (cierre)", desc: "Confirmación formal de proveedor 100% operativo en la cadena.", resp: "Compras Corporativas" },
      ]},
    ],
  },

  /* ── Sub-proceso 5: Planogramación ──────────────────────── */
  {
    id: "planograma", numero: 5, nombre: "Planogramación", color: "verde", version: "2.0", estado: "En Revisión",
    proposito: "Describir el flujo para la recepción de planillas comerciales, la elaboración técnica de planogramas y su posterior carga en sistemas de datos maestros, asegurando que la exhibición de productos responda a la estrategia comercial definida.",
    alcance_inicio: "Entrega de planilla por parte del Encargado Comercial.",
    alcance_fin: "Implementación del planograma en tiendas.",
    contexto_preceden: "Definición Comercial · Ideación de Producto.",
    contexto_suceden: "Creación de Datos Maestros · Alta de SKU en sistema.",
    actores: [
      { rol: "Encargado Comercial (PM)", area: "Comercial", resp: "Entrega planilla comercial inicial y valida ajustes de diseño." },
      { rol: "Focal Planograma", area: "Planograma", resp: "Recepción, revisión y elaboración técnica de planograma y Product List." },
      { rol: "Analista Datos Maestros", area: "Datos Maestros", resp: "Carga de la planilla de planograma en los sistemas centrales." },
      { rol: "Focal Supply Chain", area: "Abastecimiento", resp: "Gestión del envío y flujo de productos hacia tiendas." },
      { rol: "Equipo de Tiendas", area: "Tiendas", resp: "Recepción e implementación física del planograma en sala." },
    ],
    raci: {
      cols: ["Comercial Analista", "Comercial PM", "Focal Planograma", "Datos Maestros", "Focal Abastec.", "Tiendas"],
      filas: [
        { act: "Entrega de Surtido (Hito de Inicio)", vals: ["R", "A", "I", "", "", ""] },
        { act: "Recepcionar solicitud planograma", vals: ["", "", "R", "", "", ""] },
        { act: "Planogramar productos según surtido y mix", vals: ["C", "", "R", "", "", ""] },
        { act: "Generar propuesta de Product List", vals: ["", "", "R", "", "", ""] },
        { act: "¿Aprueba Validación Planograma?", vals: ["R", "A", "C", "", "", ""] },
        { act: "Carga del Surtido en Sistemas", vals: ["", "", "", "R", "", ""] },
        { act: "Confirmación de Carga completa y notificación", vals: ["", "", "", "R", "", ""] },
        { act: "Ejecución del Subproceso de Abastecimiento", vals: ["", "", "", "", "R", ""] },
        { act: "Implementación de Planogramas en Tienda", vals: ["", "", "", "", "", "R"] },
      ],
    },
    fases: [
      { nombre: "Recepción y elaboración de planograma", pasos: [
        { n: 1, act: "Entrega de Surtido (Hito de Inicio)", desc: "El área comercial entrega oficialmente la planilla inicial y define el mix estratégico de productos.", resp: "Encargado Comercial" },
        { n: 2, act: "Recepcionar solicitud planograma", desc: "Planograma ingresa formalmente la planilla y la solicitud de surtido.", resp: "Focal Planograma" },
        { n: 3, act: "¿Falta Información inicial?", desc: "Si SÍ → devolver solicitud. Si NO → planogramar.", resp: "Focal Planograma", decision: true },
        { n: 4, act: "REPROCESO: Devolver solicitud (falta de info)", desc: "Se rechaza y devuelve la solicitud al comercial para completar información.", resp: "Focal Planograma" },
        { n: 5, act: "Planogramar productos según surtido y mix", desc: "Análisis de espacio, cubicación de SKUs y diseño gráfico de la exhibición en góndola.", resp: "Focal Planograma" },
        { n: 6, act: "Generar propuesta de Product List", desc: "Listado técnico definitivo de artículos con sus ubicaciones exactas en el mueble.", resp: "Focal Planograma" },
        { n: 7, act: "¿Aprueba Validación Planograma?", desc: "Si NO → corrección de diseño. Si SÍ → generar planilla de carga.", resp: "Encargado Comercial", decision: true },
        { n: 8, act: "REPROCESO: Corrección de diseño por rechazo", desc: "El diseño es devuelto a planogramación para corregir la distribución.", resp: "Focal Planograma" },
        { n: 9, act: "Generar planilla de carga", desc: "Con validación comercial conforme, se genera la planilla técnica definitiva de carga.", resp: "Encargado Comercial" },
      ]},
      { nombre: "Carga en sistemas y ejecución en tienda", pasos: [
        { n: 10, act: "Carga del Surtido en Sistemas", desc: "Ingreso técnico de la planilla en los sistemas centrales de datos maestros.", resp: "Analista Datos Maestros" },
        { n: 11, act: "¿SKU con información completa?", desc: "Si NO → solicita información por SKU incompleto. Si SÍ → confirmación de carga.", resp: "Analista Datos Maestros", decision: true },
        { n: 14, act: "Confirmación de Carga completa y notificación", desc: "Ratificación de que el 100% de la carga está firme y emisión de cierre técnico.", resp: "Analista Datos Maestros" },
        { n: 15, act: "Ejecución del Subproceso de Abastecimiento", desc: "Se coordina envío de inventarios iniciales y órdenes de compra automáticas.", resp: "Focal Supply Chain" },
        { n: 16, act: "¿Corresponde a Salidas de Producto?", desc: "Si SÍ → Sub-proceso Salida y Sub-proceso Modificación Abastecimiento.", resp: "Encargado Comercial", decision: true },
        { n: 17, act: "Informar a las Tiendas", desc: "Emisión del comunicado oficial corporativo con directrices e inicio de la nueva carga.", resp: "Encargado Comercial" },
        { n: 18, act: "Subir diseños de planogramas a GOAP", desc: "Carga definitiva de layouts gráficos en la Guía Operativa de Actualización de Productos (GOAP).", resp: "Focal Planograma" },
        { n: 19, act: "Implementación de Planogramas en Tienda", desc: "El personal de los locales ejecuta el acomodo físico de las mercaderías en la góndola.", resp: "Equipo de Tiendas" },
        { n: 20, act: "¿Falla en implementación? (Quiebre/Error)", desc: "Si NO → fin exitoso. Si SÍ → ruta de reproceso por falla en tienda.", resp: "Equipo de Tiendas", decision: true },
      ]},
    ],
  },

  /* ── Sub-proceso 7: Condiciones de Surtido ──────────────── */
  {
    id: "surtido", numero: 7, nombre: "Condiciones de Surtido", color: "azul", version: "2.0", estado: "En Revisión",
    proposito: "Gestionar las modificaciones relacionadas con la condición de surtido de SKUs activos, incluyendo la salida formal de un producto del catálogo cuando se cumplen las condiciones requeridas. El proceso es ejecutado íntegramente por DDMM.",
    alcance_inicio: "Recepción de un ticket de solicitud de modificación de condición de surtido.",
    alcance_fin: "Cierre del ticket una vez aplicados todos los cambios y notificadas las partes. También termina si el ticket es rechazado.",
    contexto_preceden: "Gestión y Modificación de Producto.",
    contexto_suceden: "Término del proceso de Datos Maestros.",
    actores: [
      { rol: "Analista Datos Maestros", area: "DDMM", resp: "Ejecuta la totalidad del subproceso: recepciona, valida, aplica cambios en SAP y plataforma de navegación, notifica y cierra el ticket." },
      { rol: "Encargado Comercial", area: "Comercial", resp: "Origina la solicitud mediante un ticket. Recibe notificación de salida o rechazo según el resultado." },
    ],
    raci: {
      cols: ["DDMM Analista", "DDMM Jefe", "Encargado Comercial"],
      filas: [
        { act: "Recepcionar solicitud ticket", vals: ["R", "A", "I"] },
        { act: "Revisar antecedentes del ticket", vals: ["R", "A", ""] },
        { act: "Verificar completitud y datos correctos", vals: ["R", "A", ""] },
        { act: "Rechazar ticket con comentario (datos incompletos)", vals: ["R", "A", "I"] },
        { act: "Verificar si es modificación para salida de producto", vals: ["R", "A", ""] },
        { act: "Carga de información (si no es salida)", vals: ["R", "A", ""] },
        { act: "Verificar si corresponde a bloqueo de venta", vals: ["R", "A", ""] },
        { act: "Verificar si stock es 0", vals: ["R", "A", ""] },
        { act: "Rechazar ticket (stock no es 0)", vals: ["R", "A", "I"] },
        { act: "Aplicar bloqueo Z2 en SAP", vals: ["R", "A", ""] },
        { act: "Eliminar visibilidad en navegación", vals: ["R", "A", ""] },
        { act: "Quitar vigencia del producto", vals: ["R", "A", ""] },
        { act: "Quitar SKU de catálogo", vals: ["R", "A", ""] },
        { act: "Notificar salida y cerrar ticket", vals: ["R", "A", "I"] },
      ],
    },
    fases: [
      { nombre: "Recepción y validación", pasos: [
        { n: 1, act: "Recepcionar solicitud ticket", desc: "DDMM recepciona el ticket de solicitud de modificación de condición de surtido.", resp: "Datos Maestros" },
        { n: 2, act: "Revisar antecedentes", desc: "Revisión del historial del SKU, tipo de solicitud y contexto de la modificación.", resp: "Datos Maestros" },
        { n: 3, act: "Verificar completitud y datos correctos", desc: "Si NO → rechazar ticket con comentario. Si SÍ → continuar.", resp: "Datos Maestros", decision: true },
        { n: "3a", act: "Rechazo Ticket + Comentario (datos incompletos)", desc: "Se rechaza el ticket dejando un comentario explicativo. El proceso termina aquí.", resp: "Datos Maestros" },
      ]},
      { nombre: "Evaluación del tipo de cambio", pasos: [
        { n: 4, act: "Verificar si es modificación para salida de producto", desc: "Si NO → carga de información. Si SÍ → verificar bloqueo de venta.", resp: "Datos Maestros", decision: true },
        { n: "4a", act: "Carga de información (si no es salida de producto)", desc: "DDMM realiza la carga de información correspondiente.", resp: "Datos Maestros" },
        { n: 5, act: "Verificar si corresponde a bloqueo de venta", desc: "Si NO → carga de información. Si SÍ → verificar stock.", resp: "Datos Maestros", decision: true },
        { n: "5a", act: "Cargar bloqueo", desc: "Carga del bloqueo de venta correspondiente en sistema.", resp: "Datos Maestros" },
        { n: 6, act: "Verificar si stock es 0", desc: "Si SÍ → flujo de salida. Si NO → ¿tiene ajuste excepcional de inventario? Si no tiene → rechazo.", resp: "Datos Maestros", decision: true },
        { n: "6a", act: "Rechazo Ticket + Comentario (stock no es 0)", desc: "Si no hay ajuste excepcional de Inventario, se rechaza el ticket. El proceso termina.", resp: "Datos Maestros" },
      ]},
      { nombre: "Salida del producto", pasos: [
        { n: 7, act: "Aplicar bloqueo Z2 en SAP", desc: "Con stock en 0 confirmado, se aplica el bloqueo Z2 impidiendo nuevas transacciones de venta.", resp: "Datos Maestros" },
        { n: 8, act: "Eliminar visibilidad en navegación", desc: "Se elimina la visibilidad del SKU en la plataforma de navegación y catálogo digital.", resp: "Datos Maestros" },
        { n: 9, act: "Quitar vigencia del producto", desc: "El SKU queda marcado como producto no activo en el sistema.", resp: "Datos Maestros" },
        { n: 10, act: "Quitar SKU de catálogo", desc: "El SKU se retira del catálogo activo de productos.", resp: "Datos Maestros" },
        { n: 11, act: "Notificar salida", desc: "Notificación de salida formal al Encargado Comercial.", resp: "Datos Maestros" },
        { n: 12, act: "Cerrar ticket (cierre del proceso)", desc: "Cierre formal del ticket en el sistema. Evento de cierre del subproceso.", resp: "Datos Maestros" },
      ]},
    ],
  },

  /* ── Sub-proceso 8: Full (Campañas) ─────────────────────── */
  {
    id: "full", numero: 8, nombre: "Campañas Full Copec", color: "azul", version: "2.0", estado: "En Revisión",
    proposito: "Gestionar la solicitud, revisión y aprobación de campañas promocionales del programa Full Copec, garantizando que pasen por las instancias de validación comercial, operacional, técnica y contable antes de ser comunicadas a los canales.",
    alcance_inicio: "Identificación de una necesidad comercial de incluir un producto en la planilla Full.",
    alcance_fin: "Emisión del comunicado de campaña aprobada a Administración Propia y/o Franquicias, o rechazo formal en cualquier etapa.",
    contexto_preceden: "Ideación del Producto.",
    contexto_suceden: "Ciclo de Vida del Producto.",
    actores: [
      { rol: "Representante Comercial (PM)", area: "Comercial", resp: "Inicia el proceso. Evalúa si el producto requiere planilla Full y solicita campaña si aplica." },
      { rol: "Comité Fidelidad", area: "DDMM", resp: "Revisa la solicitud, verifica cumplimiento de indicadores y aprueba o rechaza." },
      { rol: "Analista / Ejecutivo DDMM", area: "DDMM", resp: "Verifica funcionamiento mecánico del SKU y habilita para QA." },
      { rol: "Responsable FULL Copec", area: "FULL Copec", resp: "Revisa la solicitud en su totalidad y carga QA para prueba." },
      { rol: "Responsable TI", area: "TI", resp: "Realiza revisión técnica y confirma pruebas aprobadas en sistema QA." },
      { rol: "Responsable Contabilidad", area: "Contabilidad", resp: "Revisa ajuste contable y verifica cuadratura de apuntes contables." },
      { rol: "Franquicias", area: "Franquicias", resp: "Recibe comunicado de campaña aprobada al cierre del proceso." },
    ],
    raci: {
      cols: ["Comercial", "Comité Fidelidad", "DDMM", "FULL Copec", "TI", "Contabilidad", "Franquicias"],
      filas: [
        { act: "Evaluar necesidad de campaña (¿Solicita?)", vals: ["R", "A", "I", "C", "I", "I", ""] },
        { act: "Rellenar Formulario Workflow", vals: ["R", "A", "I", "", "", "", "I"] },
        { act: "Revisar completitud de campaña e indicadores", vals: ["C", "I", "R", "", "", "", ""] },
        { act: "Validación SKU e Habilitación en QA", vals: ["C", "I", "I", "R", "", "", ""] },
        { act: "Gestión operativa y carga de fidelizador", vals: ["I", "I", "C", "C", "R", "", ""] },
        { act: "Revisión técnica y pruebas en sistema QA", vals: ["I", "I", "I", "C", "C", "R", ""] },
        { act: "Ajustes contables y conciliación financiera", vals: ["I", "I", "I", "I", "C", "C", "R"] },
        { act: "Verificar aprobación y selección de canal", vals: ["I", "I", "R", "", "", "", ""] },
        { act: "Emisión de Comunicado GOAP (Adm. Propia)", vals: ["R", "A", "I", "C", "", "", ""] },
        { act: "Emisión de Comunicado Franquicias", vals: ["I", "I", "A", "", "", "", "R"] },
      ],
    },
    fases: [
      { nombre: "Evaluación y solicitud", pasos: [
        { n: 1, act: "¿Solicita Campaña?", desc: "Si NO → fin temprano. Si SÍ → rellenar formulario en Workflow.", resp: "Encargado Comercial", decision: true },
        { n: 2, act: "Rellenar Formulario Workflow", desc: "Registro formal de todos los datos de la campaña en la plataforma de Workflow.", resp: "Encargado Comercial" },
      ]},
      { nombre: "Revisiones y validaciones", pasos: [
        { n: 3, act: "Revisión Comité / ¿Completitud campaña?", desc: "Si incompleta → rechazar en Workflow. Si OK → carril de DDMM.", resp: "Comité Fidelidad", decision: true },
        { n: 4, act: "Validación SKU de DDMM e Habilitación para QA", desc: "Si NO habilitado → rechazar. Si SÍ → gestión operativa FULL Copec.", resp: "DDMM", decision: true },
        { n: 5, act: "Gestión Full Copec / ¿Cargado fidelizador?", desc: "Si NO → rechazar en Workflow. Si SÍ → revisión técnica TI.", resp: "FULL Copec", decision: true },
        { n: 6, act: "Revisión TI: Sistema QA / ¿Pruebas exitosas?", desc: "Si NO → adjuntar evidencia de error y rechazar. Si SÍ → revisión contable.", resp: "TI", decision: true },
        { n: 7, act: "Ajustes y Conciliación Financiera (Contabilidad)", desc: "Si NO cuadran los apuntes → rechazar. Si SÍ → aprobación definitiva.", resp: "Contabilidad", decision: true },
      ]},
      { nombre: "Comunicación a canales", pasos: [
        { n: 8, act: "Verifica aprobación y Selección de Canal", desc: "Comité Fidelidad ratifica vigencia y define estrategia: Adm. Propia o Franquicias.", resp: "Comité Fidelidad" },
        { n: "9A", act: "Comunicado GOAP (Administración Propia)", desc: "Si aplica para locales Adm. Propia: el comercial emite el masivo interno GOAP.", resp: "Encargado Comercial / Comité Fidelidad" },
        { n: "9B", act: "Gestión y Difusión en Franquicias", desc: "Si involucra Franquicias: Comité completa formulario y Franquicias emite comunicado.", resp: "Comité Fidelidad / Franquicias" },
      ]},
    ],
  },

  /* ── Sub-proceso 9: TCT ──────────────────────────────────── */
  {
    id: "tct", numero: 9, nombre: "TCT – Tabla de Contenido de Tiendas", color: "naranja", version: "2.0", estado: "En Revisión",
    proposito: "Actualizar y mantener la Tabla de Contenido de Tiendas (TCT) con la información de SKUs que conforman el surtido, asegurando que los datos estén correctamente cargados y validados antes de su comunicación formal a las tiendas vía GOAP.",
    alcance_inicio: "Identificación de un SKU que requiere ser incorporado o actualizado en la TCT.",
    alcance_fin: "Envío del GOAP a tiendas por parte del Equipo Comercial, una vez verificada la información cargada.",
    contexto_preceden: "Ideación del Producto.",
    contexto_suceden: "Cambios y modificaciones SKU.",
    actores: [
      { rol: "Equipo Comercial", area: "Comercial / DDMM", resp: "Identifica el SKU, define categorías, completa y envía la Planilla TCT, verifica la información cargada y envía el GOAP a tiendas." },
      { rol: "Focal TCT Copec", area: "Por definir", resp: "Recepciona la información enviada por el Equipo Comercial, carga en el sistema TCT y realiza la validación técnica." },
    ],
    raci: {
      cols: ["Comercial Analista", "Comercial Jefe", "TCT Analista", "TCT Jefe"],
      filas: [
        { act: "Identificar SKU", vals: ["R", "A", "I", ""] },
        { act: "Definir categorías del SKU", vals: ["R", "A", "C", ""] },
        { act: "Completar Planilla TCT", vals: ["R", "A", "", ""] },
        { act: "Enviar Planilla al Focal TCT Copec", vals: ["R", "A", "I", ""] },
        { act: "Recepcionar información (Planilla TCT)", vals: ["I", "", "R", "A"] },
        { act: "Cargar información en sistema", vals: ["", "", "R", "A"] },
        { act: "Validación técnica de información cargada", vals: ["I", "", "R", "A"] },
        { act: "Verificación de información cargada (Comercial)", vals: ["R", "A", "C", ""] },
        { act: "Envío GOAP a tiendas", vals: ["R", "A", "", ""] },
      ],
    },
    fases: [
      { nombre: "Identificación y preparación", pasos: [
        { n: 1, act: "Identificar SKU", desc: "El Equipo Comercial identifica el SKU a incorporar o actualizar en la TCT.", resp: "Equipo Comercial" },
        { n: 2, act: "Definir categorías", desc: "Definición de las categorías correspondientes al SKU identificado.", resp: "Equipo Comercial" },
        { n: 3, act: "Completar y enviar Planilla TCT", desc: "Se completa la Planilla TCT con la información del SKU y sus categorías, y se envía al Focal TCT Copec.", resp: "Equipo Comercial" },
      ]},
      { nombre: "Carga, validación y comunicación", pasos: [
        { n: 4, act: "Recepcionar información", desc: "El Focal TCT Copec recepciona la Planilla de Actualización enviada.", resp: "Focal TCT Copec" },
        { n: 5, act: "Cargar información", desc: "El Focal TCT Copec carga la información del SKU en el sistema TCT.", resp: "Focal TCT Copec" },
        { n: 6, act: "Validación técnica de información cargada", desc: "El Focal TCT Copec valida que la información sea correcta y completa, y retorna el resultado.", resp: "Focal TCT Copec" },
        { n: 7, act: "Verificación de información cargada", desc: "El Equipo Comercial verifica la información desde la perspectiva comercial.", resp: "Equipo Comercial" },
        { n: 8, act: "Envío GOAP a tiendas", desc: "Envío formal del GOAP a las tiendas con la información verificada.", resp: "Equipo Comercial" },
      ]},
    ],
  },

  /* ── Sub-proceso 10: Modificación Abastecimiento ────────── */
  {
    id: "abastecimiento", numero: 10, nombre: "Gestión y Modificación de Abastecimiento", color: "morado", version: "2.0", estado: "En Revisión",
    proposito: "Gestionar los cambios en el modelo de distribución de un producto (Centralizado ↔ Directo), alineando a todas las áreas involucradas frente a variaciones en condiciones de venta o stock, asegurando la continuidad operativa en tiendas.",
    alcance_inicio: "Detección de caída en ventas, minstock bajo, o necesidad de cambio de modelo de distribución.",
    alcance_fin: "Recepción de productos en tiendas bajo el nuevo modelo y comunicado GOAP enviado.",
    contexto_preceden: "Subproceso de Condiciones de Abastecimiento · Seguimiento Comercial de Productos Vigentes.",
    contexto_suceden: "Seguimiento de ventas post-modificación.",
    actores: [
      { rol: "Encargado Comercial / Tiendas", area: "Comercial / Tiendas", resp: "Activan el proceso al levantar el requerimiento de cambio. Son los receptores finales del nuevo modelo." },
      { rol: "Equipo de Abastecimiento", area: "Abastecimiento (Dueño del Proceso)", resp: "Eje central del flujo: analiza factibilidad, audita in-stock, define la estrategia Directo/CD, selecciona distribuidor y emite órdenes de compra." },
      { rol: "Distribuidor (Key Logistics / Blue Express)", area: "Operadores Logísticos Externos (CD)", resp: "Alta y setup técnico de SKUs y fichas de proveedores en sus plataformas. Recepcionan pedidos y validan inventario." },
      { rol: "Datos Maestros (DDMM)", area: "DDMM", resp: "Recepcionan solicitudes formalizadas y realizan carga masiva en SAP (Libros de Pedidos, Fuentes de Aprovisionamiento y Bloqueos)." },
      { rol: "Proveedor Externo", area: "Socio Comercial", resp: "En ruta Directa, despacha y distribuye los productos directamente a tiendas sin Centro de Distribución." },
    ],
    raci: {
      cols: ["Comercial/Tiendas", "Abastecimiento", "Distribuidor (KL/Blue)", "DDMM"],
      filas: [
        { act: "Revisar solicitudes de cambio de distribución", vals: ["I", "R", "", ""] },
        { act: "Revisar indicadores de instock", vals: ["I", "R", "", ""] },
        { act: "Determinar tipo de abastecimiento (Directo / CD)", vals: ["C", "R", "", ""] },
        { act: "Gestión propia proveedor (Ruta Directa)", vals: ["I", "A", "", ""] },
        { act: "Determinar tipo de distribuidor (KL / Blue)", vals: ["", "R", "C", ""] },
        { act: "Crear SKU y Proveedor en Distribuidor", vals: ["", "I", "R", ""] },
        { act: "Enviar notificación de aprobación en CD", vals: ["I", "", "R", ""] },
        { act: "Cargar pedido a proveedor y evaluar Bloqueo", vals: ["", "R", "", ""] },
        { act: "Notificar solicitud de cambios a DDMM", vals: ["", "R", "", "I"] },
        { act: "Recibir solicitud y cargar parámetros en SAP", vals: ["I", "", "", "R"] },
        { act: "Ejecutar empujes a tiendas", vals: ["I", "R", "", ""] },
      ],
    },
    fases: [
      { nombre: "Evaluación y definición del modelo", pasos: [
        { n: 0, act: "Trigger / Inicio", desc: "Se activa cuando se reciben solicitudes de cambio en el modelo de distribución de un SKU desde Comercial o Tiendas.", resp: "Comercial / Tiendas" },
        { n: 1, act: "Revisar solicitudes de cambio de distribución", desc: "Recepción y análisis técnico del requerimiento inicial de cambio para evaluar prefactibilidad.", resp: "Equipo de Abastecimiento" },
        { n: 2, act: "Revisar indicadores de instock", desc: "Auditoría de KPIs de disponibilidad de stock para fundamentar técnicamente el cambio.", resp: "Equipo de Abastecimiento" },
        { n: 3, act: "¿Tipo de abastecimiento? (Decisión)", desc: "¿Migra a modelo Directo (entrega del proveedor) o a Centro de Distribución (CD)?", resp: "Equipo de Abastecimiento", decision: true },
        { n: "4D", act: "Gestión propia proveedor (Ruta Directa)", desc: "Si Directo → el proveedor asume autonomía total del despacho bajo condiciones pactadas.", resp: "Proveedor / Abastecimiento" },
        { n: "4CD", act: "¿Tipo de distribuidor? (Decisión)", desc: "Si CD → seleccionar operador: Key Logistics (KL) o Blue Express.", resp: "Equipo de Abastecimiento", decision: true },
      ]},
      { nombre: "Alta en distribuidor", pasos: [
        { n: "5KL", act: "Manda a crear SKU (Ruta KL)", desc: "Solicitud formal de alta y configuración técnica bajo parámetros de Key Logistics.", resp: "Abastecimiento" },
        { n: "6KL", act: "Crea SKU (Distribuidor KL)", desc: "Alta sistémica del SKU en plataformas KL. SLA estimado: 7 días.", resp: "Distribuidor (Key Logistics)" },
        { n: "7KL", act: "Crear proveedor (Distribuidor KL)", desc: "Homologación y registro del proveedor en plataformas KL. SLA: 7 días a 2 semanas.", resp: "Distribuidor (Key Logistics)" },
        { n: "5B", act: "Manda a crear SKU (Ruta Blue)", desc: "Solicitud formal de alta bajo parámetros de Blue Express.", resp: "Abastecimiento" },
        { n: "6B", act: "Crear SKU (Distribuidor Blue)", desc: "Alta del SKU en sistemas operacionales de Blue Express.", resp: "Distribuidor (Blue Express)" },
        { n: "7B", act: "Crea proveedor (Distribuidor Blue)", desc: "Registro y vinculación del proveedor en Blue Express. SLA estimado: 2 días.", resp: "Distribuidor (Blue Express)" },
        { n: 8, act: "Enviar notificación de aprobación en CD", desc: "El operador logístico confirma que el setup del producto/proveedor está aprobado para operar.", resp: "Distribuidor (KL / Blue Express)" },
      ]},
      { nombre: "Carga en SAP y empujes a tiendas", pasos: [
        { n: 9, act: "Cargar pedido a proveedor", desc: "Con la aprobación del CD, Abastecimiento carga la orden de compra inicial.", resp: "Equipo de Abastecimiento" },
        { n: 10, act: "¿Requiere Bloqueo de Compra?", desc: "Si SÍ → solicitar Libro de Pedido, Fuentes de Aprovisionamiento y Bloqueo. Si NO → omitir bloqueos.", resp: "Abastecimiento", decision: true },
        { n: 11, act: "Notifica a DDMM", desc: "Envío formal del expediente de aprovisionamiento estructurado hacia Datos Maestros.", resp: "Equipo de Abastecimiento" },
        { n: 12, act: "Recibir solicitud", desc: "DDMM acusa recibo y valida que la documentación cumpla estándares de carga.", resp: "Datos Maestros (DDMM)" },
        { n: 13, act: "Cargar datos solicitados", desc: "Parametrización y carga definitiva de datos de aprovisionamiento en SAP, activando el nuevo flujo sistémico.", resp: "Datos Maestros (DDMM)" },
        { n: 14, act: "Solicitar pedido a proveedor", desc: "Abastecimiento emite la orden de compra al proveedor o distribuidor.", resp: "Abastecimiento" },
        { n: 15, act: "Recibir pedido en bodegas CD", desc: "El distribuidor recepciona físicamente los pedidos del proveedor.", resp: "Distribuidor" },
        { n: 16, act: "Confirmar stock en CD", desc: "Auditoría y validación de inventario en bodega.", resp: "Abastecimiento / Distribuidor" },
        { n: 17, act: "Ejecutar empujes a tiendas", desc: "Abastecimiento despacha los productos hacia la red de tiendas bajo el nuevo modelo.", resp: "Equipo de Abastecimiento" },
      ]},
    ],
  },
];

/* Datos heredados (se usan en otros lugares del código si los hubiera) */
const CVP_ETAPAS = [
  {
    nombre: "Evaluación comercial e ideación",
    resumen: "¿El producto es viable? ¿Ya existe? ¿El proveedor está registrado?",
    pasos: [
      { n: 1, act: "Ideación de Producto", desc: "Se evalúa comercialmente si es conveniente incorporar el producto analizando márgenes, costos y factibilidad logística.", resp: "Encargado Comercial", decision: false },
      { n: 2, act: "¿Producto nuevo?", desc: "Si el SKU ya existe → flujo de modificaciones o bajas. Si es nuevo → continúa al siguiente paso.", resp: "Encargado Comercial", decision: true },
      { n: 3, act: "¿Existe proveedor?", desc: "Si ya está registrado en el maestro → define alcance del SKU. Si no → gestionar registro del proveedor primero.", resp: "Encargado Comercial", decision: true },
    ]
  },
  {
    nombre: "Gestión de Proveedor nuevo",
    resumen: "Solo si el proveedor no existe. Compras Corporativas audita y aprueba.",
    pasos: [
      { n: 4, act: "Solicitar creación de proveedor", desc: "Comercial recopila documentación comercial, tributaria y legal del nuevo proveedor.", resp: "Encargado Comercial", decision: false },
      { n: 5, act: "Enviar a validación", desc: "Envío formal del expediente a Compras Corporativas para auditoría y aprobación.", resp: "Encargado Comercial", decision: false },
      { n: 6, act: "Evaluación y Gestión de Proveedor", desc: "Compras Corporativas audita documentos y evalúa riesgos. Si cumple → alta definitiva. Si no → declinación y notificación.", resp: "Compras Corporativas", decision: true },
      { n: 7, act: "Notificar Encargado Comercial", desc: "Aviso confirmando que el código del proveedor está activo en los sistemas para continuar con el SKU.", resp: "Compras Corporativas", decision: false },
    ]
  },
  {
    nombre: "Definición y envío de solicitud",
    resumen: "Se define el tipo de SKU (Parcial o Total) y se envía a DDMM.",
    pasos: [
      { n: 8, act: "¿Requiere solo código de SKU?", desc: "Parcial: para importaciones o compras anticipadas. Total: para venta directa en tienda.", resp: "Encargado Comercial", decision: true },
      { n: 9, act: "Generar creación parcial (solo código SKU)", desc: "Llenado de datos mínimos (identidad, marca, jerarquía) para emitir código de barras preliminar de forma ágil.", resp: "Encargado Comercial", decision: false },
      { n: 10, act: "Generar Gran Planilla (SKU Total)", desc: "Confección de la planilla completa: dimensiones, empaque, logística, precios, costos y flujos de transporte.", resp: "Encargado Comercial", decision: false },
      { n: 11, act: "Solicitud Franquicia (camino alternativo)", desc: "Locales franquiciados solicitan alta de productos específicos directamente a través de su portal técnico.", resp: "Franquiciado", decision: false },
      { n: 12, act: "Enviar solicitud a DDMM", desc: "Despacho formal del requerimiento (planilla técnica o ticket de franquicia) hacia la mesa de Datos Maestros.", resp: "Encargado Comercial / Franquiciado", decision: false },
    ]
  },
  {
    nombre: "Procesamiento en SAP por DDMM",
    resumen: "DDMM recibe, valida y carga todas las condiciones del SKU en SAP.",
    pasos: [
      { n: 13, act: "Recepción y control de calidad", desc: "Ingreso en la Ticketera DDMM. Se revisa consistencia, duplicados y campos vacíos.", resp: "Datos Maestros (DDMM)", decision: false },
      { n: 14, act: "¿Información completa?", desc: "Si hay inconsistencias → se notifica al solicitante para corregir sin rechazar el trámite. Si está conforme → carga de datos.", resp: "DDMM / Solicitante", decision: true },
      { n: 15, act: "Carga información básica SKU en SAP", desc: "Registro inicial: descripción técnica, clasificación y dimensiones de empaque.", resp: "Datos Maestros (DDMM)", decision: false },
      { n: 16, act: "Cargar condiciones de venta", desc: "PVP, esquemas de impuestos y asignación por clústeres de tiendas.", resp: "Datos Maestros (DDMM)", decision: false },
      { n: 17, act: "Cargar condiciones de compra", desc: "Registro Info en SAP: proveedor, costo neto unitario y lead times.", resp: "Datos Maestros (DDMM)", decision: false },
      { n: 18, act: "Cargar condiciones de abastecimiento", desc: "Libro Pedido en SAP: flujo (Directo/Centralizado/Mixto) y parámetros de reposición automática.", resp: "Datos Maestros (DDMM)", decision: false },
      { n: 19, act: "Cargar condiciones de surtido", desc: "Catalogación del SKU en la matriz de tiendas autorizadas y bloqueos por sucursal.", resp: "Datos Maestros (DDMM)", decision: false },
    ]
  },
  {
    nombre: "Sincronización y cierre",
    resumen: "El SKU se propaga a SCKUBA y tiendas. Se confirma habilitación a todas las áreas.",
    pasos: [
      { n: 20, act: "Sincronización SAP → SCKUBA", desc: "El sistema propaga automáticamente la información a SCKUBA y a las terminales de tiendas.", resp: "Datos Maestros (DDMM)", decision: false },
      { n: 21, act: "Informar a solicitante + célula de cierre", desc: "Cierre del ticket y alerta automatizada confirmando a todas las áreas que el SKU está 100% activo y transaccional.", resp: "Datos Maestros (DDMM)", decision: false },
      { n: 22, act: "¿Requiere planograma?", desc: "Si necesita ordenamiento en góndola → sub-proceso Planogramación. Si no → complementación final.", resp: "Encargado Comercial", decision: true },
      { n: 23, act: "Sub-proceso Planogramación", desc: "Análisis de cubicaje, dimensiones y layout. Entrega del Product List con ubicación física y digital.", resp: "Focal de Planograma", decision: false },
      { n: 24, act: "Completar configuración SKU", desc: "Carga final de atributos comerciales pendientes y reinyección al ciclo operativo de la compañía.", resp: "Encargado Comercial", decision: false },
    ]
  },
];

const CVP_ACTORES = [
  { rol: "Representante Comercial (PM / Analista)", area: "Comercial", resp: "Evalúa viabilidad, define navegación, completa planilla de solicitud y genera ticket a DDMM." },
  { rol: "Analista focal de Datos Maestros", area: "DDMM", resp: "Control de calidad de solicitud, creación de SKU en SAP, carga de condiciones y sincronización en SCKUBA." },
  { rol: "Focal de Abastecimiento", area: "Supply Chain", resp: "Coordinación de condiciones logísticas, flujo de abastecimiento y reposición automática." },
  { rol: "Focal de Planograma", area: "Planograma", resp: "Validación de planograma y cubicación del producto." },
  { rol: "Marketing", area: "Marketing", resp: "Provisión de imagen obligatoria cuando aplica navegación digital." },
  { rol: "Franquicias", area: "Franquicias", resp: "Solicita alta de productos específicos para su operación a través del portal técnico." },
  { rol: "Compras Corporativas", area: "Compras Corporativas", resp: "Lidera la validación de proveedores nuevos: recepciona expediente, audita documentos, aprueba y notifica a DDMM y al solicitante." },
];

const CVP_RACI = [
  { act: "Evaluar viabilidad del producto", vals: ["R", "I", "C", "C", "C"] },
  { act: "Definir navegación y catalogación", vals: ["R", "I", "I", "I", "C"] },
  { act: "Completar planilla técnica", vals: ["R", "I", "I", "I", "I"] },
  { act: "Creación básica SKU en SAP", vals: ["I", "R", "I", "I", "I"] },
  { act: "Carga de condiciones (Compra/Venta/Abastecimiento/Surtido)", vals: ["I", "R", "C", "I", "I"] },
  { act: "Sincronización SAP → SCKUBA", vals: ["I", "R", "I", "I", "I"] },
  { act: "Provisión de imagen (si aplica)", vals: ["C", "I", "I", "I", "R"] },
  { act: "Validación de planograma", vals: ["A", "I", "I", "R", "I"] },
  { act: "Comunicación de cierre a áreas", vals: ["R", "I", "I", "I", "I"] },
  { act: "Escalamiento de bloqueos estructurales", vals: ["R", "C", "C", "C", "I"] },
];

/* ============================================================
   VISTA CVP: Ciclo de Vida del Producto — multi sub-proceso
   ============================================================ */
const CVP_COLOR = {
  verde:  { border: "#34c759", bg: "#f0faf2", badge: "#e8f5e9", text: "#1b5e20" },
  azul:   { border: "#5b8dee", bg: "#f0f5ff", badge: "#e8eeff", text: "#2a4a9e" },
  naranja:{ border: "#ff9500", bg: "#fff8f0", badge: "#fff3e0", text: "#bf360c" },
  rojo:   { border: "#ff3b30", bg: "#fff5f5", badge: "#ffebee", text: "#b71c1c" },
  morado: { border: "#af52de", bg: "#f9f0ff", badge: "#f3e5f5", text: "#4a148c" },
};

function VistaCVP() {
  const [selectedId, setSelectedId] = useState("mapa");
  const [openSecs, setOpenSecs] = useState({ proposito: true });

  const sp = CVP_SUBPROCESOS.find(s => s.id === selectedId) || CVP_SUBPROCESOS[0];
  const col = CVP_COLOR[sp.color] || CVP_COLOR.azul;

  const toggleSec = (k) => setOpenSecs(prev => ({ ...prev, [k]: !prev[k] }));

  const raciVal = (v) => {
    if (!v) return <td key={v} />;
    const cls = { R: "cvp2-r", A: "cvp2-a", C: "cvp2-c", I: "cvp2-i" }[v] || "";
    return <td className={cls}>{v}</td>;
  };

  return (
    <main className="sol-wrap">
      <div className="sol-head">
        <div>
          <h1 className="sol-title">Ciclo de Vida del Producto</h1>
          <p className="sol-sub">Proceso estandarizado para la creación, modificación y baja de SKUs. Área dueña: DDMM · Versión 2.0 · En Revisión</p>
        </div>
      </div>

      {/* ── Tabs horizontales de sub-procesos (desktop) / select (mobile) ── */}
      <div className="cvp2-tabs-wrap">
        {/* Desktop: tabs */}
        <nav className="cvp2-tabs">
          {CVP_SUBPROCESOS.map(s => {
            const c = CVP_COLOR[s.color] || CVP_COLOR.azul;
            const active = s.id === selectedId;
            return (
              <button
                key={s.id}
                className={"cvp2-tab-item" + (active ? " cvp2-tab-active" : "")}
                style={active ? { borderBottomColor: c.border, color: c.border } : {}}
                onClick={() => { setSelectedId(s.id); setOpenSecs({ proposito: true }); }}
              >
                <span className="cvp2-tab-num" style={active ? { background: c.border, color: "#fff" } : {}}>
                  {s.numero !== undefined ? s.numero : "★"}
                </span>
                <span className="cvp2-tab-name">{s.nombre}</span>
              </button>
            );
          })}
        </nav>
        {/* Mobile: select nativo */}
        <select
          className="cvp2-tabs-select"
          value={selectedId}
          onChange={e => { setSelectedId(e.target.value); setOpenSecs({ proposito: true }); }}
        >
          {CVP_SUBPROCESOS.map(s => (
            <option key={s.id} value={s.id}>
              {s.numero !== undefined ? `${s.numero}. ` : "★  "}{s.nombre}
            </option>
          ))}
        </select>
      </div>

        {/* ── Contenido principal ── */}
        <div className="cvp2-main">
          {/* Cabecera del sub-proceso */}
          <div className="cvp2-header" style={{ borderTopColor: col.border, background: col.bg }}>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                {sp.numero !== undefined && (
                  <span className="cvp2-sp-num" style={{ background: col.border }}>SP {sp.numero}</span>
                )}
                <span className="cvp2-estado">{sp.estado}</span>
                <span className="cvp2-version">v{sp.version}</span>
              </div>
              <h2 className="cvp2-title">{sp.nombre}</h2>
            </div>
          </div>

          {/* Propósito y Alcance */}
          <div className="cvp2-sec">
            <button className="cvp2-sec-btn" onClick={() => toggleSec("proposito")}>
              <span><BookOpen size={15} style={{ marginRight: 6, verticalAlign: "middle" }} />Propósito y Alcance</span>
              {openSecs.proposito ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
            </button>
            {openSecs.proposito && (
              <div className="cvp2-sec-body fade-in">
                <p className="cvp2-proposito">{sp.proposito}</p>
                <div className="cvp2-alcance-grid">
                  <div className="cvp2-alcance-item"><strong>Inicio (trigger):</strong> {sp.alcance_inicio}</div>
                  <div className="cvp2-alcance-item"><strong>Fin del proceso:</strong> {sp.alcance_fin}</div>
                  {sp.contexto_preceden && <div className="cvp2-alcance-item"><strong>Preceden:</strong> {sp.contexto_preceden}</div>}
                  {sp.contexto_suceden && <div className="cvp2-alcance-item"><strong>Suceden:</strong> {sp.contexto_suceden}</div>}
                </div>
              </div>
            )}
          </div>

          {/* Actores */}
          <div className="cvp2-sec">
            <button className="cvp2-sec-btn" onClick={() => toggleSec("actores")}>
              <span><User size={15} style={{ marginRight: 6, verticalAlign: "middle" }} />Actores e Integrantes</span>
              {openSecs.actores ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
            </button>
            {openSecs.actores && (
              <div className="cvp2-sec-body fade-in">
                <div className="cvp2-actores-grid">
                  {sp.actores.map((a, i) => (
                    <div key={i} className="cvp2-actor-card">
                      <div className="cvp2-actor-rol">{a.rol}</div>
                      <div className="cvp2-actor-area" style={{ color: col.border }}>{a.area}</div>
                      <div className="cvp2-actor-resp">{a.resp}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* RACI */}
          <div className="cvp2-sec">
            <button className="cvp2-sec-btn" onClick={() => toggleSec("raci")}>
              <span><ClipboardList size={15} style={{ marginRight: 6, verticalAlign: "middle" }} />Matriz RACI</span>
              {openSecs.raci ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
            </button>
            {openSecs.raci && (
              <div className="cvp2-sec-body fade-in">
                <div className="cvp2-raci-scroll">
                  <table className="cvp2-raci-tbl">
                    <thead>
                      <tr>
                        <th>Actividad / Decisión</th>
                        {sp.raci.cols.map((c, i) => <th key={i}>{c}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {sp.raci.filas.map((f, i) => (
                        <tr key={i}>
                          <td>{f.act}</td>
                          {f.vals.map((v, j) => <td key={j} className={v ? "cvp2-" + v.toLowerCase() : ""}>{v}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <p className="cvp2-raci-legend">R = Responsable · A = Accountable · C = Consultado · I = Informado</p>
              </div>
            )}
          </div>

          {/* Pasos del proceso */}
          <div className="cvp2-sec">
            <button className="cvp2-sec-btn" onClick={() => toggleSec("pasos")}>
              <span><ArrowRight size={15} style={{ marginRight: 6, verticalAlign: "middle" }} />Flujo del Proceso · Pasos</span>
              {openSecs.pasos ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
            </button>
            {openSecs.pasos && (
              <div className="cvp2-sec-body fade-in">
                {sp.fases.map((fase, fi) => (
                  <div key={fi} className="cvp2-fase">
                    <div className="cvp2-fase-title" style={{ borderLeftColor: col.border }}>{fase.nombre}</div>
                    <div className="cvp2-pasos-list">
                      {fase.pasos.map((p, pi) => (
                        <div key={pi} className={"cvp2-paso" + (p.decision ? " cvp2-paso-decision" : "")}>
                          <span className="cvp2-paso-n" style={p.decision ? { background: col.border, color: "#fff" } : {}}>{p.n}</span>
                          <div className="cvp2-paso-cont">
                            <strong>{p.act}</strong>
                            <span>{p.desc}</span>
                            <em>{p.resp}</em>
                          </div>
                          {p.decision && <span className="cvp2-decision-tag" style={{ background: col.badge, color: col.text }}>⬦ Decisión</span>}
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
    </main>
  );
}


function VistaClusters({ clusters, onChange }) {
  const [pegado, setPegado] = useState("");
  const [copiado, setCopiado] = useState(false);
  const [nombre, setNombre] = useState("");
  const [listaCluster, setListaCluster] = useState("");
  const [editando, setEditando] = useState(null); // nombre original en edición
  const [msj, setMsj] = useState(null);

  /* Convertidor: lista vertical (o con tabs/; ) -> códigos */
  const codigos = React.useMemo(() => {
    return pegado.split(/[\n\r\t;]+/).map(code).map(norm).filter(Boolean)
      .filter((v, i, a) => a.indexOf(v) === i);
  }, [pegado]);
  const linea = codigos.join(",");
  const validos = codigos.filter(c => MAESTROS.centros[c]?.nombre);
  const invalidos = codigos.filter(c => !MAESTROS.centros[c]?.nombre);

  const copiar = async () => {
    try { await navigator.clipboard.writeText(linea); }
    catch {
      const ta = document.createElement("textarea");
      ta.value = linea; document.body.appendChild(ta); ta.select();
      document.execCommand("copy"); document.body.removeChild(ta);
    }
    setCopiado(true); setTimeout(() => setCopiado(false), 1800);
  };

  /* Mantenedor */
  const centrosForm = React.useMemo(
    () => listaCluster.split(",").map(code).filter(Boolean).filter((v, i, a) => a.indexOf(v) === i),
    [listaCluster]
  );
  const invalidosForm = centrosForm.filter(c => !MAESTROS.centros[c]?.nombre);

  const guardarCluster = () => {
    const n = norm(nombre);
    if (n.length < 2) { setMsj("El nombre debe tener al menos 2 caracteres."); return; }
    if (/^\d+$/.test(n)) { setMsj("El nombre no puede ser solo números (se confundiría con un código de centro)."); return; }
    if (!centrosForm.length) { setMsj("Agrega al menos un centro al cluster."); return; }
    const existeOtro = Object.keys(clusters).some(k => k.toLowerCase() === n.toLowerCase() && k !== editando);
    if (existeOtro) { setMsj(`Ya tienes un cluster llamado "${n}".`); return; }
    const next = { ...clusters };
    if (editando && editando !== n) delete next[editando];
    next[n] = centrosForm;
    onChange(next);
    setNombre(""); setListaCluster(""); setEditando(null);
    setMsj(`Cluster "${n}" guardado con ${centrosForm.length} centros.`);
  };

  const editar = n => { setNombre(n); setListaCluster(clusters[n].join(",")); setEditando(n); setMsj(null); window.scrollTo({ top: 0, behavior: "smooth" }); };
  const eliminar = n => { const next = { ...clusters }; delete next[n]; onChange(next); if (editando === n) { setNombre(""); setListaCluster(""); setEditando(null); } };
  const copiarCluster = async n => {
    try { await navigator.clipboard.writeText(clusters[n].join(",")); } catch {}
  };

  const nombres = Object.keys(clusters).sort((a, b) => a.localeCompare(b));

  return (
    <main className="sol-wrap">
      <div className="sol-head">
        <div>
          <h1 className="sol-title">Gestor de locales</h1>
          <p className="sol-sub">Convierte listas de locales y guarda tus clusters. En las planillas, escribe el nombre del cluster en la columna Centro y se expandirá automáticamente al validar.</p>
        </div>
      </div>

      <div className="clu-grid">
        {/* CONVERTIDOR */}
        <div className="clu-card">
          <div className="clu-card-head"><ArrowLeftRight size={17} /><h2>Convertidor de lista</h2></div>
          <p className="clu-help">Pega aquí tu columna de locales (uno por línea, tal como la copias de Excel) y la transformamos en una lista separada por comas.</p>
          <textarea
            className="clu-textarea"
            value={pegado}
            onChange={e => setPegado(e.target.value)}
            placeholder={"2003\n2549\n2626"}
            rows={7}
          />
          {codigos.length > 0 && (
            <>
              <div className="clu-resultado">
                <span className="clu-linea">{linea}</span>
                <button className="btn-soft" onClick={copiar}>{copiado ? <><Check size={14} /> Copiado</> : <><Copy size={14} /> Copiar</>}</button>
              </div>
              <div className="clu-conteo">
                <span className="t-ok"><CheckCircle2 size={13} /> {validos.length} válidos</span>
                {invalidos.length > 0 && <span className="t-err"><XCircle size={13} /> {invalidos.length} no existen: {invalidos.slice(0, 8).join(", ")}{invalidos.length > 8 && "…"}</span>}
              </div>
              <button className="btn-ghost" onClick={() => { setListaCluster(linea); setMsj(null); }}>
                <Layers size={14} /> Usar esta lista para crear un cluster →
              </button>
            </>
          )}
        </div>

        {/* MANTENEDOR */}
        <div className="clu-card">
          <div className="clu-card-head"><Layers size={17} /><h2>Mis clusters</h2></div>
          <p className="clu-help">Guarda listas de centros con un nombre. Son personales: solo tú los ves y puedes usarlos en cualquier planilla.</p>
          <div className="clu-form">
            <input className="clu-input" value={nombre} placeholder='Nombre del cluster (ej: "Prontos RM")'
              onChange={e => { setNombre(e.target.value); setMsj(null); }} />
            <textarea className="clu-textarea sm" rows={3} value={listaCluster}
              placeholder="Centros separados por coma: 2003,2549,2626"
              onChange={e => { setListaCluster(e.target.value); setMsj(null); }} />
            {centrosForm.length > 0 && (
              <div className="clu-conteo">
                <span>{centrosForm.length} centros</span>
                {invalidosForm.length > 0 && <span className="t-err"><XCircle size={13} /> {invalidosForm.length} no existen: {invalidosForm.slice(0, 6).join(", ")}{invalidosForm.length > 6 && "…"}</span>}
              </div>
            )}
            <div className="clu-form-btns">
              <button className="btn-validar" onClick={guardarCluster}><Save size={14} /> {editando ? "Guardar cambios" : "Crear cluster"}</button>
              {editando && <button className="btn-ghost" onClick={() => { setNombre(""); setListaCluster(""); setEditando(null); setMsj(null); }}>Cancelar edición</button>}
            </div>
            {msj && <div className="aviso-info">{msj}</div>}
          </div>

          {nombres.length === 0 ? (
            <div className="clu-vacio">Aún no tienes clusters guardados.</div>
          ) : (
            <div className="clu-lista">
              {nombres.map(n => {
                const cs = clusters[n];
                const malos = cs.filter(c => !MAESTROS.centros[code(c)]?.nombre).length;
                return (
                  <div key={n} className="clu-item">
                    <div className="clu-item-info">
                      <strong>{n}</strong>
                      <span>{cs.length} centros · {cs.slice(0, 6).join(", ")}{cs.length > 6 && "…"}{malos > 0 && <em className="t-err"> · {malos} no existen</em>}</span>
                    </div>
                    <div className="fila-btns">
                      <button className="dz-clear dup" onClick={() => copiarCluster(n)} title="Copiar lista" aria-label="Copiar lista"><Copy size={13} /></button>
                      <button className="dz-clear dup" onClick={() => editar(n)} title="Editar" aria-label="Editar"><Pencil size={13} /></button>
                      <button className="dz-clear" onClick={() => eliminar(n)} title="Eliminar" aria-label="Eliminar"><Trash2 size={13} /></button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      <div className="clu-tip">
        <ShieldCheck size={15} />
        <span>En las planillas puedes combinar: <code>Prontos RM,2745</code> usará todos los centros del cluster más el 2745. Los duplicados dentro de la celda se eliminan solos. También funciona en los Excel importados.</span>
      </div>
    </main>
  );
}

/* ============================================================
   VISTA ADMIN: gestión de usuarios y roles
   ============================================================ */
/* ============================================================
   VISTA MAESTROS: mantenedor de tablas maestras por tabla
   ============================================================ */
const TABLAS_CONFIG = [
  {
    id: "centros",
    nombre: "Centros",
    desc: "Locales con organización de ventas. Columnas esperadas: Codigo centro, Nombre centro, Sociedad.",
    icon: Database,
    tabla: "centros",
    conflicto: "codigo",
    columnas: [
      { key: "codigo",    label: "Codigo centro",  col: 0, req: true },
      { key: "nombre",    label: "Nombre centro",   col: 1, req: true },
      { key: "org_venta", label: "Sociedad",        col: 2, req: false, def: "BP01" },
    ],
    templateEjemplo: [["2003", "Pronto Pargua", "BP01"], ["2549", "Pronto Rosario Norte", "BP01"]],
  },
  {
    id: "skus",
    nombre: "SKUs",
    desc: "Materiales con unidad de venta y condición. Columnas esperadas: Codigo, Nombre, Unidad Venta, Unidad Condicion.",
    icon: Boxes,
    tabla: "skus",
    conflicto: "codigo",
    columnas: [
      { key: "codigo",           label: "Codigo",          col: 0, req: true },
      { key: "nombre",           label: "Nombre",          col: 1, req: true },
      { key: "unidad_venta",     label: "Unidad Venta",    col: 2, req: false, def: "UN" },
      { key: "unidad_condicion", label: "Unidad Condicion",col: 3, req: false, def: "UN" },
    ],
    templateEjemplo: [["32", "COCA-COLA ZERO 250C", "UN", "UN"], ["145873", "MATERIAL EJEMPLO", "CJ", "UN"]],
  },
  {
    id: "proveedores",
    nombre: "Proveedores",
    desc: "Proveedores registrados en SAP. Columnas esperadas: Codigo, Nombre.",
    icon: Truck,
    tabla: "proveedores",
    conflicto: "codigo",
    columnas: [
      { key: "codigo", label: "Codigo", col: 0, req: true },
      { key: "nombre", label: "Nombre", col: 1, req: true },
    ],
    templateEjemplo: [["100000000", "Comercial Puerto Chico Ltda"], ["100000108", "Embotelladora Andina S.A"]],
  },
  {
    id: "grupos_articulo",
    nombre: "Grupos Artículo",
    desc: "Grupos de artículo SAP. Columnas esperadas: Codigo, Nombre.",
    icon: Tag,
    tabla: "grupos_articulo",
    conflicto: "codigo",
    columnas: [
      { key: "codigo", label: "Codigo", col: 0, req: true },
      { key: "nombre", label: "Nombre", col: 1, req: true },
    ],
    templateEjemplo: [["ALIM", "Alimentos"], ["BEBE", "Bebidas"]],
  },
  {
    id: "materiales_modelo",
    nombre: "Materiales Modelo",
    desc: "Materiales modelo para catalogación. Columnas esperadas: Codigo, Nombre, Descripcion.",
    icon: FileText,
    tabla: "materiales_modelo",
    conflicto: "codigo",
    columnas: [
      { key: "codigo",      label: "Codigo",      col: 0, req: true },
      { key: "nombre",      label: "Nombre",      col: 1, req: true },
      { key: "descripcion", label: "Descripcion", col: 2, req: false, def: "" },
    ],
    templateEjemplo: [["MAT001", "Material Ejemplo", "Descripcion del material"], ["MAT002", "Otro Material", ""]],
  },
  {
    id: "marcas",
    nombre: "Marcas",
    desc: "Marcas de productos. Columnas esperadas: Codigo, Nombre.",
    icon: Tag,
    tabla: "marcas",
    conflicto: "codigo",
    columnas: [
      { key: "codigo", label: "Codigo", col: 0, req: true },
      { key: "nombre", label: "Nombre", col: 1, req: true },
    ],
    templateEjemplo: [["COCA", "Coca-Cola"], ["NESTLE", "Nestlé"]],
  },
  {
    id: "fabricantes",
    nombre: "Fabricantes",
    desc: "Fabricantes de productos. Columnas esperadas: Codigo, Nombre.",
    icon: Truck,
    tabla: "fabricantes",
    conflicto: "codigo",
    columnas: [
      { key: "codigo", label: "Codigo", col: 0, req: true },
      { key: "nombre", label: "Nombre", col: 1, req: true },
    ],
    templateEjemplo: [["FAB001", "Nestlé Chile S.A."], ["FAB002", "Coca-Cola Andina"]],
  },
];

function VistaMaestros() {
  const [tabActiva, setTabActiva] = useState("centros");
  const fileRefs = Object.fromEntries(TABLAS_CONFIG.map(t => [t.id, useRef()]));
  const [previews, setPreviews] = useState({});
  const [subiendo, setSubiendo] = useState({});
  const [resultado, setResultado] = useState({});

  const leerExcel = (cfg, file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const sh = wb.Sheets[wb.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });
        if (rawRows.length < 2) {
          setPreviews(p => ({ ...p, [cfg.id]: { error: "El archivo no contiene datos.", fileName: file.name } }));
          return;
        }
        const rows = [];
        for (let i = 1; i < rawRows.length; i++) {
          const raw = rawRows[i];
          const r = {};
          cfg.columnas.forEach(c => { r[c.key] = norm(raw[c.col]) || c.def || ""; });
          // ignorar filas completamente vacías
          if (cfg.columnas.filter(c => c.req).every(c => !r[c.key])) continue;
          const errs = cfg.columnas.filter(c => c.req && !r[c.key]).map(c => `${c.label} vacío`);
          rows.push({ data: r, errores: errs, ok: errs.length === 0 });
        }
        setPreviews(p => ({ ...p, [cfg.id]: { rows, fileName: file.name } }));
        setResultado(r => ({ ...r, [cfg.id]: null }));
      } catch {
        setPreviews(p => ({ ...p, [cfg.id]: { error: "No se pudo leer el archivo.", fileName: file.name } }));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const subirTabla = async (cfg) => {
    const prev = previews[cfg.id];
    if (!prev?.rows) return;
    const filas_ok = prev.rows.filter(r => r.ok);
    if (!filas_ok.length) return;
    setSubiendo(s => ({ ...s, [cfg.id]: true }));
    const { error } = await supabase.from(cfg.tabla).upsert(filas_ok.map(r => r.data), { onConflict: cfg.conflicto });
    setSubiendo(s => ({ ...s, [cfg.id]: false }));
    if (error) {
      setResultado(r => ({ ...r, [cfg.id]: { tipo: "error", texto: error.message } }));
    } else {
      setResultado(r => ({ ...r, [cfg.id]: { tipo: "ok", texto: `${filas_ok.length} registros actualizados. Recarga el portal para usar la nueva base.` } }));
      setPreviews(p => ({ ...p, [cfg.id]: null }));
    }
  };

  const descargarPlantilla = (cfg) => {
    const headers = cfg.columnas.map(c => c.label);
    const aoa = [headers, ...cfg.templateEjemplo];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = headers.map(h => ({ wch: Math.max(h.length + 4, 16) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, cfg.nombre);
    XLSX.writeFile(wb, `plantilla_${cfg.id}.xlsx`);
  };

  const cfg = TABLAS_CONFIG.find(t => t.id === tabActiva);
  const prev = previews[tabActiva];
  const res = resultado[tabActiva];
  const filasOk = prev?.rows?.filter(r => r.ok).length ?? 0;
  const filasErr = prev?.rows?.filter(r => !r.ok).length ?? 0;

  return (
    <main className="sol-wrap">
      <div className="sol-head">
        <div>
          <h1 className="sol-title">Bases de Datos Maestras</h1>
          <p className="sol-sub">Actualiza cada tabla de forma independiente subiendo un Excel. Los registros existentes se actualizan, los nuevos se agregan.</p>
        </div>
      </div>

      {/* Tabs de tablas */}
      <div className="mto-tabs">
        {TABLAS_CONFIG.map(t => {
          const Icon = t.icon;
          return (
            <button key={t.id} className={"mto-tab" + (tabActiva === t.id ? " on" : "")} onClick={() => setTabActiva(t.id)}>
              <Icon size={16} /> {t.nombre}
              {previews[t.id]?.rows && (
                <span className={"mto-badge " + (previews[t.id].rows.some(r => !r.ok) ? "err" : "ok")}>
                  {previews[t.id].rows.length}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* Panel activo */}
      <div className="mto-panel">
        <p className="mto-desc">{cfg.desc}</p>

        {/* Acciones */}
        <div className="mto-actions">
          <button className="btn-soft" onClick={() => descargarPlantilla(cfg)}>
            <Download size={15} /> Descargar plantilla
          </button>
          <input type="file" accept=".xlsx,.xls" hidden ref={fileRefs[cfg.id]}
            onChange={e => { leerExcel(cfg, e.target.files[0]); e.target.value = ""; }} />
          <button className="btn-validar" onClick={() => fileRefs[cfg.id].current?.click()}>
            <Upload size={15} /> Subir Excel
          </button>
        </div>

        {/* Error de lectura */}
        {prev?.error && (
          <div className="mto-msg err"><XCircle size={15} /> {prev.error}</div>
        )}

        {/* Preview */}
        {prev?.rows && (
          <>
            <div className="mto-summary">
              <span className="mto-file"><strong>{prev.fileName}</strong></span>
              <span className="mto-cnt ok"><CheckCircle2 size={14} /> {filasOk} válidas</span>
              {filasErr > 0 && <span className="mto-cnt err"><XCircle size={14} /> {filasErr} con errores</span>}
            </div>

            <div className="mto-table-wrap">
              <table className="mto-table">
                <thead>
                  <tr>
                    <th></th>
                    {cfg.columnas.map(c => <th key={c.key}>{c.label}</th>)}
                    <th>Observaciones</th>
                  </tr>
                </thead>
                <tbody>
                  {prev.rows.slice(0, 50).map((r, i) => (
                    <tr key={i} className={r.ok ? "" : "fila-err"}>
                      <td>{r.ok ? <CheckCircle2 size={14} color="#248a3d" /> : <XCircle size={14} color="#c2271c" />}</td>
                      {cfg.columnas.map(c => <td key={c.key}>{r.data[c.key] || <span className="mto-vacio">—</span>}</td>)}
                      <td>{r.errores.join(", ") || ""}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {prev.rows.length > 50 && (
                <p className="mto-mas">Mostrando primeras 50 filas de {prev.rows.length} total.</p>
              )}
            </div>

            <div className="mto-actions" style={{ marginTop: 16 }}>
              <button
                className="btn-enviar"
                disabled={filasOk === 0 || subiendo[cfg.id]}
                onClick={() => subirTabla(cfg)}
              >
                {subiendo[cfg.id] ? "Actualizando..." : `Confirmar y actualizar ${filasOk} registros en Supabase`}
              </button>
            </div>
          </>
        )}

        {/* Resultado de subida */}
        {res && (
          <div className={"mto-msg " + res.tipo}>
            {res.tipo === "ok" ? <CheckCircle2 size={15} /> : <XCircle size={15} />} {res.texto}
          </div>
        )}
      </div>
    </main>
  );
}

/* ============================================================
   CREACIÓN DE NUEVOS SKUs
   ============================================================ */
function SkuFieldLabel({ children }) {
  return <div style={{ fontSize: 12, color: "#636366", marginBottom: 4 }}>{children}</div>;
}
function SkuSectionLabel({ children }) {
  return <div style={{ fontSize: 12, fontWeight: 700, color: "#5b8dee", textTransform: "uppercase", letterSpacing: "0.05em", margin: "18px 0 10px" }}>{children}</div>;
}

// Columnas del template Excel de importación de SKUs
const SKU_EXCEL_COLS = [
  { key: "nombre",            label: "Nombre",              col: 0 },
  { key: "descripcion",       label: "Descripcion",         col: 1 },
  { key: "tipo",              label: "Tipo",                col: 2 },
  { key: "grupo_articulo",    label: "Grupo Articulo",      col: 3 },
  { key: "material_modelo",   label: "Material Modelo",     col: 4 },
  { key: "marca",             label: "Marca",               col: 5 },
  { key: "fabricante",        label: "Fabricante",          col: 6 },
  { key: "flag_planograma",   label: "Flag Planograma",     col: 7 },
  { key: "flag_navegacion",   label: "Flag Navegacion",     col: 8 },
  { key: "_p_unidad",         label: "Primaria Unidad",     col: 9 },
  { key: "_p_alto",           label: "Primaria Alto",       col: 10 },
  { key: "_p_largo",          label: "Primaria Largo",      col: 11 },
  { key: "_p_ancho",          label: "Primaria Ancho",      col: 12 },
  { key: "_p_peso_neto",      label: "Primaria Peso Neto",  col: 13 },
  { key: "_p_peso_bruto",     label: "Primaria Peso Bruto", col: 14 },
  { key: "factor_conversion", label: "Factor Conversion",   col: 15 },
  { key: "_s_unidad",         label: "Secundaria Unidad",   col: 16 },
  { key: "_s_alto",           label: "Secundaria Alto",     col: 17 },
  { key: "_s_largo",          label: "Secundaria Largo",    col: 18 },
  { key: "_s_ancho",          label: "Secundaria Ancho",    col: 19 },
  { key: "_s_peso_neto",      label: "Secundaria Peso Neto",col: 20 },
  { key: "_s_peso_bruto",     label: "Secundaria Peso Bruto",col: 21 },
];

function VistaCreacionSKU({ perfil, session, vista, setVista }) {
  const [skus, setSkus] = useState([emptySkuRow()]);
  const [expandedRow, setExpandedRow] = useState(null);
  const [colGroup, setColGroup] = useState("basicos");
  const solicitante = perfil?.nombre || "";
  const [enviando, setEnviando] = useState(false);
  const [enviado, setEnviado] = useState(null);
  const [cat, setCat] = useState({ grupos: [], modelos: [], marcas: [], fabricantes: [] });
  const importRef = useRef();

  useEffect(() => {
    (async () => {
      const [{ data: grupos }, { data: modelos }, { data: marcas }, { data: fabricantes }] = await Promise.all([
        supabase.from("grupos_articulo").select("codigo,nombre"),
        supabase.from("materiales_modelo").select("codigo,nombre"),
        supabase.from("marcas").select("codigo,nombre"),
        supabase.from("fabricantes").select("codigo,nombre"),
      ]);
      setCat({ grupos: grupos || [], modelos: modelos || [], marcas: marcas || [], fabricantes: fabricantes || [] });
    })();
  }, []);

  const importarExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const dataRows = raw.slice(1).filter(r => r[0]);
      const nuevos = dataRows.slice(0, 80).map(row => {
        const s = emptySkuRow();
        SKU_EXCEL_COLS.forEach(c => {
          const val = String(row[c.col] ?? "").trim();
          if (c.key.startsWith("_p_")) s.dim_primaria[c.key.slice(3)] = val;
          else if (c.key.startsWith("_s_")) s.dim_secundaria[c.key.slice(3)] = val;
          else if (c.key === "flag_planograma" || c.key === "flag_navegacion")
            s[c.key] = val.toUpperCase() === "SI" || val === "1" || val.toUpperCase() === "TRUE";
          else s[c.key] = val;
        });
        return s;
      });
      if (nuevos.length) setSkus(nuevos);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const descargarTemplate = () => {
    const headers = SKU_EXCEL_COLS.map(c => c.label);
    const ejemplo = ["Agua Mineral 500ml", "Agua mineral sin gas botella PET", "Retail", "BEB", "MAT001", "NESTLE", "FAB001", "SI", "SI", "22", "7", "7", "0.5", "0.55", "24", "8", "8", "12", "13.2"];
    const ws = XLSX.utils.aoa_to_sheet([headers, ejemplo]);
    ws["!cols"] = headers.map(h => ({ wch: Math.max(h.length + 2, 14) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SKUs");
    XLSX.writeFile(wb, "template_nuevos_sku.xlsx");
  };

  const updateSku    = (id, k, v)      => setSkus(s => s.map(x => x._id === id ? { ...x, [k]: v } : x));
  const updateSkuDim = (id, dim, k, v) => setSkus(s => s.map(x => x._id === id ? { ...x, [dim]: { ...x[dim], [k]: v } } : x));
  const addSku       = () => { if (skus.length >= 80) return; setSkus(s => [...s, emptySkuRow()]); };
  const removeSku    = (id) => setSkus(s => s.filter(x => x._id !== id));

  const addFotos   = (id, tipo, files) => setSkus(s => s.map(x => x._id !== id ? x : { ...x, [`fotos_${tipo}`]: [...x[`fotos_${tipo}`], ...Array.from(files)] }));
  const removeFoto = (id, tipo, i)     => setSkus(s => s.map(x => x._id !== id ? x : { ...x, [`fotos_${tipo}`]: x[`fotos_${tipo}`].filter((_, j) => j !== i) }));

  const validarTodo = () => {
    let ok = true;
    const updated = skus.map(sku => {
      const errs = [];
      if (!sku.nombre.trim())   errs.push("Nombre requerido");
      if (!sku.tipo)            errs.push("Tipo requerido");
      if (!sku.grupo_articulo)  errs.push("Grupo artículo requerido");
      if (!sku.material_modelo) errs.push("Material modelo requerido");
      if (!sku.marca)           errs.push("Marca requerida");
      if (!sku.fabricante)      errs.push("Fabricante requerido");
      if (!TIPOS_SIN_DIM.has(sku.tipo)) {
        ["dim_primaria", "dim_secundaria"].forEach((dim, di) => {
          DIM_CAMPOS.forEach(c => { if (!sku[dim][c.key]) errs.push(`${di === 0 ? "Primaria" : "Secundaria"}: ${c.label} requerido`); });
        });
      }
      if (!TIPOS_SIN_FOTO.has(sku.tipo)) {
        if (sku.fotos_planograma.length < 3) errs.push("Planograma: mínimo 3 fotos JPG");
        if (sku.fotos_navegacion.length === 0) errs.push("Navegación: se requiere 1 foto PNG");
      }
      if (errs.length) ok = false;
      return { ...sku, errores: errs };
    });
    setSkus(updated);
    const firstErr = updated.find(s => s.errores.length);
    if (firstErr) setExpandedRow(firstErr._id);
    return ok;
  };

  const handleEnviar = async () => {
    if (!validarTodo() || solicitante.trim().length < 3) return;
    setEnviando(true);
    const folio = "SKU-" + new Date().getFullYear() + "-" + String(Math.floor(100000 + Math.random() * 899999));
    const filas = await Promise.all(skus.map(async (sku, idx) => {
      const uploadDir = async (files, carpeta) => {
        const paths = [];
        for (const file of files) {
          const path = `${folio}/${idx}/${carpeta}/${file.name}`;
          const { error } = await supabase.storage.from("sku-fotos").upload(path, file, { upsert: true });
          if (!error) paths.push(path);
        }
        return paths;
      };
      let fp = [], fn = [];
      if (!TIPOS_SIN_FOTO.has(sku.tipo)) {
        fp = await uploadDir(sku.fotos_planograma, "planograma");
        fn = await uploadDir(sku.fotos_navegacion, "navegacion");
      }
      const { _id, errores, fotos_planograma, fotos_navegacion, ...data } = sku;
      return { ...data, fotos_planograma_paths: fp, fotos_navegacion_paths: fn };
    }));
    const { error } = await supabase.from("solicitudes").insert({
      folio,
      solicitante_id:     session?.user?.id,
      solicitante_nombre: solicitante,
      solicitante_email:  session?.user?.email,
      estado:    "Enviada",
      planillas: [{ id: "creacion_sku", nombre: "Creación de SKU", filas }],
      historial: [{ estado: "Enviada", fecha: new Date().toISOString() }],
    });
    setEnviando(false);
    if (!error) setEnviado({ folio, total: skus.length });
    else alert("Error al guardar: " + JSON.stringify(error));
  };

  if (enviado) return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil}>
      <main className="success-wrap">
        <div className="success-check"><CheckCircle2 size={64} strokeWidth={1.3} /></div>
        <h1 className="success-title">Solicitud enviada.</h1>
        <p className="success-sub">Folio <strong>{enviado.folio}</strong> · {enviado.total} SKU{enviado.total > 1 ? "s" : ""} a crear.</p>
        <button className="btn-primary" style={{ marginTop: 24 }} onClick={() => { setEnviado(null); setSkus([emptySkuRow()]); }}>Nueva solicitud</button>
      </main>
    </AppShell>
  );

  const totalFotos   = skus.filter(s => !TIPOS_SIN_FOTO.has(s.tipo)).reduce((a, s) => a + s.fotos_planograma.length + s.fotos_navegacion.length, 0);
  const skusConError = skus.filter(s => s.errores.length > 0);

  const panelDerecho = (
    <div style={{ padding: "28px 20px", display: "flex", flexDirection: "column", gap: 16 }}>
      <div>
        <div style={{ fontSize: 13, fontWeight: 700, color: "#3a3a3c", marginBottom: 12 }}>Resumen</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span style={{ color: "#636366" }}>SKUs</span>
            <strong>{skus.length} / 80</strong>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span style={{ color: "#636366" }}>Fotos adjuntas</span>
            <strong>{totalFotos}</strong>
          </div>
          {skusConError.length > 0 && (
            <div style={{ fontSize: 12, color: "#c2271c", background: "rgba(255,59,48,0.07)", borderRadius: 8, padding: "8px 10px", marginTop: 4 }}>
              {skusConError.length} SKU{skusConError.length > 1 ? "s" : ""} con errores
            </div>
          )}
        </div>
      </div>
      <div style={{ borderTop: "1px solid rgba(0,0,0,0.08)", paddingTop: 16 }}>
        <div style={{ fontSize: 12, color: "#636366", marginBottom: 6 }}>Solicitante</div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <User size={14} color="#636366" />
          <span style={{ fontSize: 13, color: solicitante ? "#3a3a3c" : "#aeaeb2", fontWeight: solicitante ? 500 : 400 }}>
            {solicitante || "Sin usuario autenticado"}
          </span>
        </div>
        <button className="btn-primary" style={{ width: "100%", marginBottom: 8 }} disabled={solicitante.trim().length < 3 || enviando} onClick={handleEnviar}>
          {enviando ? "Enviando…" : "Enviar solicitud"} <ArrowRight size={15} />
        </button>
      </div>
      <div style={{ borderTop: "1px solid rgba(0,0,0,0.08)", paddingTop: 16 }}>
        <p style={{ fontSize: 11, color: "#aeaeb2", lineHeight: 1.5 }}>Usa los grupos de columnas para navegar entre secciones. Las fotos están en el grupo "Fotos / Flags".</p>
      </div>
    </div>
  );

  const thSt = { padding: "8px 10px", textAlign: "left", fontSize: 11, fontWeight: 600, color: "#86868b", textTransform: "uppercase", letterSpacing: ".05em", borderBottom: "1px solid rgba(200,205,230,0.4)", background: "#f8f8fc", whiteSpace: "nowrap" };

  const COL_GROUPS = [
    { id: "basicos",     label: "Datos básicos" },
    { id: "marcas",      label: "Marca / Prov." },
    { id: "dimensiones", label: "Dimensiones" },
    { id: "fotos",       label: "Fotos / Flags" },
  ];

  const dimNA = (key) => <td key={key} style={{ background: "#f8f8fa" }}><span style={{ color: "#d2d2d7", fontSize: 12, padding: "7px 8px", display: "block" }}>—</span></td>;

  return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil} rightPanel={panelDerecho}>
      <main className="sol-wrap">
        <div className="bloque">
          <div className="bloque-head">
            <span className="card-icon sm"><Package size={18} strokeWidth={1.7} /></span>
            <div className="bloque-tit">
              <strong>Nuevo SKU</strong>
              <span>{skusConError.length > 0 ? `${skusConError.length} SKU${skusConError.length > 1 ? "s" : ""} con error` : "Pendiente de validación"}</span>
            </div>
            <div className="bloque-acciones">
              <button className="btn-ghost" onClick={descargarTemplate}><Download size={14} /> Plantilla</button>
              <input type="file" accept=".xlsx,.xls" hidden ref={importRef} onChange={importarExcel} />
              <button className="btn-ghost" onClick={() => importRef.current?.click()}><Upload size={14} /> Importar Excel</button>
            </div>
          </div>

          {/* Selector de grupo de columnas */}
          <div style={{ display: "flex", gap: 6, padding: "8px 12px", borderBottom: "1px solid rgba(0,0,0,0.06)", flexWrap: "wrap" }}>
            {COL_GROUPS.map(g => (
              <button key={g.id} onClick={() => setColGroup(g.id)} style={{
                fontSize: 12, padding: "4px 14px", borderRadius: 980, border: "1px solid",
                borderColor: colGroup === g.id ? "#5b8dee" : "rgba(0,0,0,0.12)",
                background: colGroup === g.id ? "#5b8dee" : "transparent",
                color: colGroup === g.id ? "#fff" : "#515154",
                cursor: "pointer", fontWeight: colGroup === g.id ? 600 : 400, transition: "all .15s",
              }}>{g.label}</button>
            ))}
          </div>

          <div className="grilla-scroll">
            <table className="grilla">
              <thead>
                <tr>
                  <th className="th-n">#</th>
                  <th style={{ minWidth: 140 }}>ID Creación</th>

                  {colGroup === "basicos" && <>
                    <th style={{ minWidth: 175 }}>Nombre</th>
                    <th style={{ minWidth: 185 }}>Descripción</th>
                    <th style={{ minWidth: 120 }}>Tipo</th>
                  </>}

                  {colGroup === "marcas" && <>
                    <th style={{ minWidth: 140 }}>Grupo Art.</th>
                    <th style={{ minWidth: 140 }}>Mat. Modelo</th>
                    <th style={{ minWidth: 140 }}>Marca</th>
                    <th style={{ minWidth: 140 }}>Fabricante</th>
                  </>}

                  {colGroup === "dimensiones" && <>
                    <th style={{ minWidth: 80, borderLeft: "2px solid rgba(91,141,238,0.25)", background: "rgba(91,141,238,0.04)", color: "#5b8dee" }}>Unid. P</th>
                    <th style={{ minWidth: 75, background: "rgba(91,141,238,0.04)" }}>Alto P</th>
                    <th style={{ minWidth: 75, background: "rgba(91,141,238,0.04)" }}>Largo P</th>
                    <th style={{ minWidth: 75, background: "rgba(91,141,238,0.04)" }}>Ancho P</th>
                    <th style={{ minWidth: 85, background: "rgba(91,141,238,0.04)" }}>P.Neto P</th>
                    <th style={{ minWidth: 85, background: "rgba(91,141,238,0.04)" }}>P.Bruto P</th>
                    <th style={{ minWidth: 90, textAlign: "center" }}>Factor Conv.</th>
                    <th style={{ minWidth: 80, borderLeft: "2px solid rgba(94,92,230,0.25)", background: "rgba(94,92,230,0.04)", color: "#5e5ce6" }}>Unid. S</th>
                    <th style={{ minWidth: 75, background: "rgba(94,92,230,0.04)" }}>Alto S</th>
                    <th style={{ minWidth: 75, background: "rgba(94,92,230,0.04)" }}>Largo S</th>
                    <th style={{ minWidth: 75, background: "rgba(94,92,230,0.04)" }}>Ancho S</th>
                    <th style={{ minWidth: 85, background: "rgba(94,92,230,0.04)" }}>P.Neto S</th>
                    <th style={{ minWidth: 85, background: "rgba(94,92,230,0.04)" }}>P.Bruto S</th>
                  </>}

                  {colGroup === "fotos" && <>
                    <th style={{ minWidth: 65, textAlign: "center" }}>Plan.</th>
                    <th style={{ minWidth: 65, textAlign: "center" }}>Nav.</th>
                    <th style={{ minWidth: 110 }}>Fotos</th>
                  </>}

                  <th className="th-x"></th>
                </tr>
              </thead>
              <tbody>
                {skus.map((sku, idx) => {
                  const sinDim   = TIPOS_SIN_DIM.has(sku.tipo);
                  const sinFoto  = TIPOS_SIN_FOTO.has(sku.tipo);
                  const hayError = sku.errores.length > 0;
                  const expanded = expandedRow === sku._id;
                  const NCOLS_NOW = colGroup === "basicos" ? 6 : colGroup === "marcas" ? 7 : colGroup === "dimensiones" ? 16 : 6;

                  return (
                    <React.Fragment key={sku._id}>
                      <tr className={hayError ? "g-err" : ""}>
                        <td className="td-n">
                          {hayError ? <AlertTriangle size={13} className="ic-err" /> : <Pencil size={12} className="ic-dim" />}
                        </td>
                        <td>
                          <span style={{ fontFamily: "monospace", fontSize: 11, background: "#f0f0f5", color: "#5e5ce6", borderRadius: 6, padding: "2px 7px", whiteSpace: "nowrap" }}>
                            {sku.id_creacion}
                          </span>
                        </td>

                        {colGroup === "basicos" && <>
                          <td><input className="celda" value={sku.nombre} onChange={e => updateSku(sku._id, "nombre", e.target.value)} placeholder="Nombre del SKU" /></td>
                          <td><input className="celda" value={sku.descripcion} onChange={e => updateSku(sku._id, "descripcion", e.target.value)} placeholder="Descripción" /></td>
                          <td>
                            <select className="celda" value={sku.tipo} onChange={e => updateSku(sku._id, "tipo", e.target.value)}>
                              <option value="">—</option>
                              {TIPOS_SKU_BASE.map(t => <option key={t}>{t}</option>)}
                            </select>
                          </td>
                        </>}

                        {colGroup === "marcas" && <>
                          <td>
                            <select className="celda" value={sku.grupo_articulo} onChange={e => updateSku(sku._id, "grupo_articulo", e.target.value)}>
                              <option value="">—</option>
                              {cat.grupos.map(g => <option key={g.codigo} value={g.codigo}>{g.codigo} — {g.nombre}</option>)}
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={sku.material_modelo} onChange={e => updateSku(sku._id, "material_modelo", e.target.value)}>
                              <option value="">—</option>
                              {cat.modelos.map(m => <option key={m.codigo} value={m.codigo}>{m.codigo} — {m.nombre}</option>)}
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={sku.marca} onChange={e => updateSku(sku._id, "marca", e.target.value)}>
                              <option value="">—</option>
                              {cat.marcas.map(m => <option key={m.codigo} value={m.codigo}>{m.codigo} — {m.nombre}</option>)}
                              <option value="MARCA NO CREADA">NO CREADA</option>
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={sku.fabricante} onChange={e => updateSku(sku._id, "fabricante", e.target.value)}>
                              <option value="">—</option>
                              {cat.fabricantes.map(f => <option key={f.codigo} value={f.codigo}>{f.codigo} — {f.nombre}</option>)}
                              <option value="FABRICANTE NO CREADO">NO CREADO</option>
                            </select>
                          </td>
                        </>}

                        {colGroup === "dimensiones" && (sinDim ? <>
                          {[...Array(13)].map((_, i) => dimNA(`d${i}`))}
                        </> : <>
                          <td style={{ borderLeft: "2px solid rgba(91,141,238,0.2)" }}>
                            <select className="celda" value={sku.dim_primaria.unidad} onChange={e => updateSkuDim(sku._id, "dim_primaria", "unidad", e.target.value)}>
                              <option value="">—</option>
                              {UNIDADES_DIM.map(u => <option key={u}>{u}</option>)}
                            </select>
                          </td>
                          {DIM_CAMPOS.map(c => (
                            <td key={c.key}>
                              <input className="celda" type="number" min="0" step="0.001" value={sku.dim_primaria[c.key]} onChange={e => updateSkuDim(sku._id, "dim_primaria", c.key, e.target.value)} placeholder="0" />
                            </td>
                          ))}
                          <td>
                            <input className="celda" type="number" min="0" step="0.001" value={sku.factor_conversion} onChange={e => updateSku(sku._id, "factor_conversion", e.target.value)} placeholder="1" />
                          </td>
                          <td style={{ borderLeft: "2px solid rgba(94,92,230,0.2)" }}>
                            <select className="celda" value={sku.dim_secundaria.unidad} onChange={e => updateSkuDim(sku._id, "dim_secundaria", "unidad", e.target.value)}>
                              <option value="">—</option>
                              {UNIDADES_DIM.map(u => <option key={u}>{u}</option>)}
                            </select>
                          </td>
                          {DIM_CAMPOS.map(c => (
                            <td key={c.key}>
                              <input className="celda" type="number" min="0" step="0.001" value={sku.dim_secundaria[c.key]} onChange={e => updateSkuDim(sku._id, "dim_secundaria", c.key, e.target.value)} placeholder="0" />
                            </td>
                          ))}
                        </>)}

                        {colGroup === "fotos" && <>
                          <td style={{ textAlign: "center" }}>
                            <input type="checkbox" checked={sku.flag_planograma} onChange={e => updateSku(sku._id, "flag_planograma", e.target.checked)} />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <input type="checkbox" checked={sku.flag_navegacion} onChange={e => updateSku(sku._id, "flag_navegacion", e.target.checked)} />
                          </td>
                          <td>
                            {sinFoto ? <span style={{ color: "#d2d2d7", fontSize: 12, padding: "7px 8px", display: "block" }}>—</span> : (
                              <button className="dz-clear dup" style={{ padding: "4px 8px", fontSize: 12, whiteSpace: "nowrap" }} onClick={() => setExpandedRow(expanded ? null : sku._id)}>
                                <Upload size={12} /> {sku.fotos_planograma.length + sku.fotos_navegacion.length > 0 ? `${sku.fotos_planograma.length + sku.fotos_navegacion.length} foto${sku.fotos_planograma.length + sku.fotos_navegacion.length > 1 ? "s" : ""}` : "Adjuntar"}
                              </button>
                            )}
                          </td>
                        </>}

                        <td className="td-x">
                          <span className="fila-btns">
                            {skus.length > 1 && (
                              <button className="dz-clear" onClick={() => removeSku(sku._id)} title="Eliminar"><Trash2 size={13} /></button>
                            )}
                          </span>
                        </td>
                      </tr>

                      {hayError && (
                        <tr className="g-err-detalle">
                          <td></td>
                          <td colSpan={NCOLS_NOW - 1}>{sku.errores.join(" · ")}</td>
                        </tr>
                      )}

                      {/* Panel fotos expandido (solo grupo fotos) */}
                      {expanded && colGroup === "fotos" && !sinFoto && (
                        <tr>
                          <td></td>
                          <td colSpan={NCOLS_NOW - 1} style={{ padding: "14px 12px 18px", background: "#f8f8fc", borderBottom: "1px solid rgba(0,0,0,0.06)" }}>
                            <div style={{ display: "flex", gap: 14 }}>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 11, color: "#636366", marginBottom: 6 }}>Planograma · JPG (mín. 3)</div>
                                <label style={{ display: "block", border: "1.5px dashed rgba(0,0,0,0.15)", borderRadius: 10, padding: "10px", cursor: "pointer", textAlign: "center", fontSize: 12, color: "#636366" }}>
                                  <Upload size={14} style={{ display: "block", margin: "0 auto 4px" }} />
                                  {sku.fotos_planograma.length > 0 ? `${sku.fotos_planograma.length} foto${sku.fotos_planograma.length > 1 ? "s" : ""}` : "Seleccionar"}
                                  <input type="file" multiple accept="image/jpeg,image/jpg" style={{ display: "none" }} onChange={e => addFotos(sku._id, "planograma", e.target.files)} />
                                </label>
                                {sku.fotos_planograma.map((f, fi) => (
                                  <div key={fi} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#636366", marginTop: 3 }}>
                                    <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
                                    <button onClick={() => removeFoto(sku._id, "planograma", fi)} style={{ background: "none", border: "none", cursor: "pointer", color: "#aeaeb2", padding: 0 }}><X size={11} /></button>
                                  </div>
                                ))}
                              </div>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 11, color: "#636366", marginBottom: 6 }}>Navegación · PNG 600×400</div>
                                <label style={{ display: "block", border: "1.5px dashed rgba(0,0,0,0.15)", borderRadius: 10, padding: "10px", cursor: "pointer", textAlign: "center", fontSize: 12, color: "#636366" }}>
                                  <Upload size={14} style={{ display: "block", margin: "0 auto 4px" }} />
                                  {sku.fotos_navegacion.length > 0 ? "1 foto" : "Seleccionar"}
                                  <input type="file" accept="image/png" style={{ display: "none" }} onChange={e => addFotos(sku._id, "navegacion", e.target.files)} />
                                </label>
                                {sku.fotos_navegacion.length > 0 && (
                                  <div style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#636366", marginTop: 3 }}>
                                    <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sku.fotos_navegacion[0].name}</span>
                                    <button onClick={() => removeFoto(sku._id, "navegacion", 0)} style={{ background: "none", border: "none", cursor: "pointer", color: "#aeaeb2", padding: 0 }}><X size={11} /></button>
                                  </div>
                                )}
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button className="btn-addrow" onClick={addSku} disabled={skus.length >= 80}>
            <Plus size={14} /> Agregar SKU
          </button>
        </div>

      </main>
    </AppShell>
  );
}

function VistaCreacionReceta({ perfil, session, vista, setVista }) {
  const [recetas, setRecetas] = useState([emptySkuRow()]);
  const [expandedRow, setExpandedRow] = useState(null);
  const [insumoModal, setInsumoModal] = useState(null);
  const [colGroup, setColGroup] = useState("basicos");
  const solicitante = perfil?.nombre || "";
  const [enviando, setEnviando] = useState(false);
  const [enviado, setEnviado] = useState(null);
  const [cat, setCat] = useState({ grupos: [], modelos: [], marcas: [], fabricantes: [] });
  const importRef = useRef();
  const insumoImportRef = useRef();

  useEffect(() => {
    (async () => {
      const [{ data: grupos }, { data: modelos }, { data: marcas }, { data: fabricantes }] = await Promise.all([
        supabase.from("grupos_articulo").select("codigo,nombre"),
        supabase.from("materiales_modelo").select("codigo,nombre"),
        supabase.from("marcas").select("codigo,nombre"),
        supabase.from("fabricantes").select("codigo,nombre"),
      ]);
      setCat({ grupos: grupos || [], modelos: modelos || [], marcas: marcas || [], fabricantes: fabricantes || [] });
    })();
  }, []);

  const updateRec    = (id, k, v) => setRecetas(s => s.map(x => x._id === id ? { ...x, [k]: v } : x));
  const addRec       = () => { if (recetas.length >= 80) return; setRecetas(s => [...s, emptySkuRow()]); };
  const removeRec    = (id) => setRecetas(s => s.filter(x => x._id !== id));

  const addInsumo    = (id)          => setRecetas(s => s.map(x => x._id !== id ? x : { ...x, insumos: [...x.insumos, { sku: "", cantidad: "", unidad: "" }] }));
  const removeInsumo = (id, i)       => setRecetas(s => s.map(x => x._id !== id ? x : { ...x, insumos: x.insumos.filter((_, j) => j !== i) }));
  const updateInsumo = (id, i, k, v) => setRecetas(s => s.map(x => {
    if (x._id !== id) return x;
    const ins = x.insumos.map((r, j) => {
      if (j !== i) return r;
      const updated = { ...r, [k]: v };
      if (k === "sku") updated.unidad = MAESTROS.skus?.[v]?.unidadVenta || "";
      return updated;
    });
    return { ...x, insumos: ins };
  }));

  const addFotos   = (id, tipo, files) => setRecetas(s => s.map(x => x._id !== id ? x : { ...x, [`fotos_${tipo}`]: [...x[`fotos_${tipo}`], ...Array.from(files)] }));
  const removeFoto = (id, tipo, i)     => setRecetas(s => s.map(x => x._id !== id ? x : { ...x, [`fotos_${tipo}`]: x[`fotos_${tipo}`].filter((_, j) => j !== i) }));

  const descargarTemplateInsumos = () => {
    const ws = XLSX.utils.aoa_to_sheet([["Codigo SKU", "Cantidad"], ["10047", "2.5"]]);
    ws["!cols"] = [{ wch: 16 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Insumos");
    XLSX.writeFile(wb, "template_insumos_receta.xlsx");
  };
  const importarInsumos = (id, e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const nuevos = raw.slice(1).filter(r => r[0]).map(r => {
        const skuCod = String(r[0]).trim();
        return { sku: skuCod, cantidad: String(r[1] ?? "").trim(), unidad: MAESTROS.skus?.[skuCod]?.unidadVenta || "" };
      });
      if (nuevos.length) setRecetas(s => s.map(x => x._id !== id ? x : { ...x, insumos: nuevos }));
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const importarExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const nuevas = raw.slice(1).filter(r => r[0]).slice(0, 80).map(row => {
        const s = emptySkuRow();
        s.nombre           = String(row[0] ?? "").trim();
        s.descripcion      = String(row[1] ?? "").trim();
        s.grupo_articulo   = String(row[2] ?? "").trim();
        s.material_modelo  = String(row[3] ?? "").trim();
        s.marca            = String(row[4] ?? "").trim();
        s.fabricante       = String(row[5] ?? "").trim();
        return s;
      });
      if (nuevas.length) setRecetas(nuevas);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };
  const descargarTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Nombre", "Descripcion", "Grupo Articulo", "Material Modelo", "Marca", "Fabricante"],
      ["Sandwich Pollo", "Sandwich pollo grillado con lechuga", "SND", "MAT001", "MARCA01", "FAB001"],
    ]);
    ws["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Recetas");
    XLSX.writeFile(wb, "template_nuevas_recetas.xlsx");
  };

  const validarTodo = () => {
    let ok = true;
    const updated = recetas.map(rec => {
      const errs = [];
      if (!rec.nombre.trim())    errs.push("Nombre requerido");
      if (!rec.grupo_articulo)   errs.push("Grupo artículo requerido");
      if (!rec.material_modelo)  errs.push("Material modelo requerido");
      if (!rec.marca)            errs.push("Marca requerida");
      if (!rec.fabricante)       errs.push("Fabricante requerido");
      if (rec.insumos.length === 0) errs.push("Debe tener al menos un insumo");
      if (errs.length) ok = false;
      return { ...rec, errores: errs };
    });
    setRecetas(updated);
    return ok;
  };

  const handleEnviar = async () => {
    if (!validarTodo() || solicitante.trim().length < 3) return;
    setEnviando(true);
    const folio = "SAP-" + new Date().getFullYear() + "-" + String(Math.floor(100000 + Math.random() * 899999));
    const filas = recetas.map(({ _id, errores, fotos_planograma, fotos_navegacion, ...data }) => data);
    const { error } = await supabase.from("solicitudes").insert({
      folio,
      solicitante_id:     session?.user?.id,
      solicitante_nombre: solicitante,
      solicitante_email:  session?.user?.email,
      estado:    "Enviada",
      planillas: [{ id: "creacion_receta", nombre: "Creación de Receta", filas }],
      historial: [{ estado: "Enviada", fecha: new Date().toISOString() }],
    });
    setEnviando(false);
    if (!error) setEnviado({ folio, total: recetas.length });
    else alert("Error al guardar: " + JSON.stringify(error));
  };

  if (enviado) return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil}>
      <main className="success-wrap">
        <div className="success-check"><CheckCircle2 size={64} strokeWidth={1.3} /></div>
        <h1 className="success-title">Solicitud enviada.</h1>
        <p className="success-sub">Folio <strong>{enviado.folio}</strong> · {enviado.total} receta{enviado.total > 1 ? "s" : ""} a crear.</p>
        <button className="btn-primary" style={{ marginTop: 24 }} onClick={() => { setEnviado(null); setRecetas([emptySkuRow()]); }}>Nueva solicitud</button>
      </main>
    </AppShell>
  );

  const recetasConError = recetas.filter(r => r.errores.length > 0);
  const recModal = insumoModal ? recetas.find(r => r._id === insumoModal) : null;

  const thSt = { padding: "8px 10px", textAlign: "left", fontSize: 11, fontWeight: 600, color: "#86868b", textTransform: "uppercase", letterSpacing: ".05em", borderBottom: "1px solid rgba(200,205,230,0.4)", background: "#f8f8fc", whiteSpace: "nowrap" };

  const COL_GROUPS = [
    { id: "basicos", label: "Datos básicos" },
    { id: "marcas",  label: "Marca / Prov." },
    { id: "fotos",   label: "Fotos / Flags" },
  ];

  const panelDerecho = (
    <div style={{ padding: "28px 20px", display: "flex", flexDirection: "column", gap: 16 }}>
      <div>
        <div style={{ fontSize: 13, fontWeight: 700, color: "#3a3a3c", marginBottom: 12 }}>Resumen</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span style={{ color: "#636366" }}>Recetas</span>
            <strong>{recetas.length} / 80</strong>
          </div>
          {recetasConError.length > 0 && (
            <div style={{ fontSize: 12, color: "#c2271c", background: "rgba(255,59,48,0.07)", borderRadius: 8, padding: "8px 10px", marginTop: 4 }}>
              {recetasConError.length} receta{recetasConError.length > 1 ? "s" : ""} con errores
            </div>
          )}
        </div>
      </div>
      <div style={{ borderTop: "1px solid rgba(0,0,0,0.08)", paddingTop: 16 }}>
        <div style={{ fontSize: 12, color: "#636366", marginBottom: 6 }}>Solicitante</div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <User size={14} color="#636366" />
          <span style={{ fontSize: 13, color: solicitante ? "#3a3a3c" : "#aeaeb2", fontWeight: solicitante ? 500 : 400 }}>
            {solicitante || "Sin usuario autenticado"}
          </span>
        </div>
        <button className="btn-primary" style={{ width: "100%", marginBottom: 8 }} disabled={solicitante.trim().length < 3 || enviando} onClick={handleEnviar}>
          {enviando ? "Enviando…" : "Enviar solicitud"} <ArrowRight size={15} />
        </button>
      </div>
    </div>
  );

  return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil} rightPanel={panelDerecho}>
      <main className="sol-wrap">
        <div className="bloque">
          <div className="bloque-head">
            <span className="card-icon sm"><Utensils size={18} strokeWidth={1.7} /></span>
            <div className="bloque-tit">
              <strong>Nueva Receta</strong>
              <span>{recetasConError.length > 0 ? `${recetasConError.length} con error` : "Pendiente de validación"}</span>
            </div>
            <div className="bloque-acciones">
              <button className="btn-ghost" onClick={descargarTemplate}><Download size={14} /> Plantilla</button>
              <input type="file" accept=".xlsx,.xls" hidden ref={importRef} onChange={importarExcel} />
              <button className="btn-ghost" onClick={() => importRef.current?.click()}><Upload size={14} /> Importar Excel</button>
            </div>
          </div>

          <div style={{ display: "flex", gap: 6, padding: "8px 12px", borderBottom: "1px solid rgba(0,0,0,0.06)", flexWrap: "wrap" }}>
            {COL_GROUPS.map(g => (
              <button key={g.id} onClick={() => setColGroup(g.id)} style={{
                fontSize: 12, padding: "4px 14px", borderRadius: 980, border: "1px solid",
                borderColor: colGroup === g.id ? "#34c759" : "rgba(0,0,0,0.12)",
                background: colGroup === g.id ? "#34c759" : "transparent",
                color: colGroup === g.id ? "#fff" : "#515154",
                cursor: "pointer", fontWeight: colGroup === g.id ? 600 : 400, transition: "all .15s",
              }}>{g.label}</button>
            ))}
          </div>

          <div className="grilla-scroll">
            <table className="grilla">
              <thead>
                <tr>
                  <th className="th-n">#</th>
                  <th style={{ minWidth: 140 }}>ID Creación</th>
                  {colGroup === "basicos" && <>
                    <th style={{ minWidth: 175 }}>Nombre</th>
                    <th style={{ minWidth: 185 }}>Descripción</th>
                    <th style={{ minWidth: 110 }}>Insumos</th>
                  </>}
                  {colGroup === "marcas" && <>
                    <th style={{ minWidth: 140 }}>Grupo Art.</th>
                    <th style={{ minWidth: 140 }}>Mat. Modelo</th>
                    <th style={{ minWidth: 140 }}>Marca</th>
                    <th style={{ minWidth: 140 }}>Fabricante</th>
                  </>}
                  {colGroup === "fotos" && <>
                    <th style={{ minWidth: 65, textAlign: "center" }}>Plan.</th>
                    <th style={{ minWidth: 65, textAlign: "center" }}>Nav.</th>
                    <th style={{ minWidth: 110 }}>Fotos</th>
                  </>}
                  <th className="th-x"></th>
                </tr>
              </thead>
              <tbody>
                {recetas.map((rec, idx) => {
                  const hayError = rec.errores.length > 0;
                  const expanded = expandedRow === rec._id;
                  const NCOLS_NOW = colGroup === "basicos" ? 6 : colGroup === "marcas" ? 7 : 6;
                  return (
                    <React.Fragment key={rec._id}>
                      <tr className={hayError ? "g-err" : ""}>
                        <td className="td-n">
                          {hayError ? <AlertTriangle size={13} className="ic-err" /> : <Pencil size={12} className="ic-dim" />}
                        </td>
                        <td>
                          <span style={{ fontFamily: "monospace", fontSize: 11, background: "#f0f0f5", color: "#5e5ce6", borderRadius: 6, padding: "2px 7px", whiteSpace: "nowrap" }}>
                            {rec.id_creacion}
                          </span>
                        </td>
                        {colGroup === "basicos" && <>
                          <td><input className="celda" value={rec.nombre} onChange={e => updateRec(rec._id, "nombre", e.target.value)} placeholder="Nombre de la receta" /></td>
                          <td><input className="celda" value={rec.descripcion} onChange={e => updateRec(rec._id, "descripcion", e.target.value)} placeholder="Descripción" /></td>
                          <td>
                            <button className="dz-clear dup" style={{ padding: "4px 8px", fontSize: 12, whiteSpace: "nowrap" }} onClick={() => setInsumoModal(rec._id)}>
                              <Layers size={12} /> {rec.insumos.length > 0 ? `${rec.insumos.length} insumo${rec.insumos.length > 1 ? "s" : ""}` : "Insumos"}
                            </button>
                          </td>
                        </>}
                        {colGroup === "marcas" && <>
                          <td>
                            <select className="celda" value={rec.grupo_articulo} onChange={e => updateRec(rec._id, "grupo_articulo", e.target.value)}>
                              <option value="">—</option>
                              {cat.grupos.map(g => <option key={g.codigo} value={g.codigo}>{g.codigo} — {g.nombre}</option>)}
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={rec.material_modelo} onChange={e => updateRec(rec._id, "material_modelo", e.target.value)}>
                              <option value="">—</option>
                              {cat.modelos.map(m => <option key={m.codigo} value={m.codigo}>{m.codigo} — {m.nombre}</option>)}
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={rec.marca} onChange={e => updateRec(rec._id, "marca", e.target.value)}>
                              <option value="">—</option>
                              {cat.marcas.map(m => <option key={m.codigo} value={m.codigo}>{m.codigo} — {m.nombre}</option>)}
                              <option value="MARCA NO CREADA">NO CREADA</option>
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={rec.fabricante} onChange={e => updateRec(rec._id, "fabricante", e.target.value)}>
                              <option value="">—</option>
                              {cat.fabricantes.map(f => <option key={f.codigo} value={f.codigo}>{f.codigo} — {f.nombre}</option>)}
                              <option value="FABRICANTE NO CREADO">NO CREADO</option>
                            </select>
                          </td>
                        </>}
                        {colGroup === "fotos" && <>
                          <td style={{ textAlign: "center" }}>
                            <input type="checkbox" checked={rec.flag_planograma} onChange={e => updateRec(rec._id, "flag_planograma", e.target.checked)} />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <input type="checkbox" checked={rec.flag_navegacion} onChange={e => updateRec(rec._id, "flag_navegacion", e.target.checked)} />
                          </td>
                          <td>
                            <button className="dz-clear dup" style={{ padding: "4px 8px", fontSize: 12, whiteSpace: "nowrap" }} onClick={() => setExpandedRow(expanded ? null : rec._id)}>
                              <Upload size={12} /> {rec.fotos_planograma.length + rec.fotos_navegacion.length > 0 ? `${rec.fotos_planograma.length + rec.fotos_navegacion.length} foto${rec.fotos_planograma.length + rec.fotos_navegacion.length > 1 ? "s" : ""}` : "Adjuntar"}
                            </button>
                          </td>
                        </>}
                        <td className="td-x">
                          <span className="fila-btns">
                            {recetas.length > 1 && (
                              <button className="dz-clear" onClick={() => removeRec(rec._id)} title="Eliminar"><Trash2 size={13} /></button>
                            )}
                          </span>
                        </td>
                      </tr>
                      {hayError && (
                        <tr className="g-err-detalle">
                          <td></td>
                          <td colSpan={NCOLS_NOW - 1}>{rec.errores.join(" · ")}</td>
                        </tr>
                      )}
                      {expanded && colGroup === "fotos" && (
                        <tr>
                          <td></td>
                          <td colSpan={NCOLS_NOW - 1} style={{ padding: "14px 12px 18px", background: "#f8f8fc", borderBottom: "1px solid rgba(0,0,0,0.06)" }}>
                            <div style={{ display: "flex", gap: 14 }}>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 11, color: "#636366", marginBottom: 6 }}>Planograma · JPG (mín. 3)</div>
                                <label style={{ display: "block", border: "1.5px dashed rgba(0,0,0,0.15)", borderRadius: 10, padding: "10px", cursor: "pointer", textAlign: "center", fontSize: 12, color: "#636366" }}>
                                  <Upload size={14} style={{ display: "block", margin: "0 auto 4px" }} />
                                  {rec.fotos_planograma.length > 0 ? `${rec.fotos_planograma.length} foto${rec.fotos_planograma.length > 1 ? "s" : ""}` : "Seleccionar"}
                                  <input type="file" multiple accept="image/jpeg,image/jpg" style={{ display: "none" }} onChange={e => addFotos(rec._id, "planograma", e.target.files)} />
                                </label>
                                {rec.fotos_planograma.map((f, fi) => (
                                  <div key={fi} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#636366", marginTop: 3 }}>
                                    <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
                                    <button onClick={() => removeFoto(rec._id, "planograma", fi)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30", padding: 0 }}><X size={12} /></button>
                                  </div>
                                ))}
                              </div>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 11, color: "#636366", marginBottom: 6 }}>Navegación · JPG (mín. 3)</div>
                                <label style={{ display: "block", border: "1.5px dashed rgba(0,0,0,0.15)", borderRadius: 10, padding: "10px", cursor: "pointer", textAlign: "center", fontSize: 12, color: "#636366" }}>
                                  <Upload size={14} style={{ display: "block", margin: "0 auto 4px" }} />
                                  {rec.fotos_navegacion.length > 0 ? `${rec.fotos_navegacion.length} foto${rec.fotos_navegacion.length > 1 ? "s" : ""}` : "Seleccionar"}
                                  <input type="file" multiple accept="image/jpeg,image/jpg" style={{ display: "none" }} onChange={e => addFotos(rec._id, "navegacion", e.target.files)} />
                                </label>
                                {rec.fotos_navegacion.map((f, fi) => (
                                  <div key={fi} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#636366", marginTop: 3 }}>
                                    <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
                                    <button onClick={() => removeFoto(rec._id, "navegacion", fi)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30", padding: 0 }}><X size={12} /></button>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button className="btn-addrow" onClick={addRec} disabled={recetas.length >= 80}>
            <Plus size={14} /> Agregar receta
          </button>
        </div>

        {/* Modal insumos */}
        {recModal && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}
               onClick={e => { if (e.target === e.currentTarget) setInsumoModal(null); }}>
            <div style={{ background: "#fff", borderRadius: 20, width: "100%", maxWidth: 660, maxHeight: "80vh", display: "flex", flexDirection: "column", overflow: "hidden", boxShadow: "0 20px 60px rgba(0,0,0,0.22)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "18px 22px", borderBottom: "1px solid rgba(0,0,0,0.07)" }}>
                <Layers size={18} color="#34c759" />
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>Insumos de la receta</div>
                  <div style={{ fontSize: 12, color: "#636366" }}>
                    {recModal.nombre || "Receta sin nombre"} · <span style={{ fontFamily: "monospace", color: "#5e5ce6" }}>{recModal.id_creacion}</span>
                  </div>
                </div>
                <button onClick={() => setInsumoModal(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "#636366", padding: 4 }}><X size={20} /></button>
              </div>
              <div style={{ flex: 1, overflow: "auto", padding: "16px 22px" }}>
                <div style={{ border: "1px solid rgba(0,0,0,0.07)", borderRadius: 12, overflow: "hidden" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 14 }}>
                    <thead>
                      <tr>
                        <th style={thSt}>Código SKU</th>
                        <th style={thSt}>Nombre</th>
                        <th style={{ ...thSt, width: 110 }}>Cantidad</th>
                        <th style={{ ...thSt, width: 80 }}>Unidad</th>
                        <th style={{ ...thSt, width: 36 }}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {recModal.insumos.length === 0 && (
                        <tr><td colSpan={5} style={{ textAlign: "center", color: "#aeaeb2", padding: "22px 10px", fontSize: 13 }}>Sin insumos aún. Agrega el primero abajo.</td></tr>
                      )}
                      {recModal.insumos.map((ins, iIdx) => {
                        const skuInfo = MAESTROS.skus?.[ins.sku];
                        return (
                          <tr key={iIdx} style={{ borderBottom: "1px solid rgba(0,0,0,0.04)" }}>
                            <td style={{ padding: "3px 6px" }}>
                              <input className="celda" value={ins.sku} onChange={e => updateInsumo(recModal._id, iIdx, "sku", e.target.value)} placeholder="Código" style={{ minWidth: 100 }} />
                            </td>
                            <td style={{ padding: "4px 10px" }}>
                              {skuInfo ? <span style={{ fontSize: 13, color: "#34c759" }}>{skuInfo.nombre}</span>
                               : ins.sku ? <span style={{ fontSize: 13, color: "#ff3b30" }}>No encontrado</span>
                               : <span style={{ fontSize: 13, color: "#c7c7cc" }}>—</span>}
                            </td>
                            <td style={{ padding: "3px 6px" }}>
                              <input className="celda" type="number" min="0" step="0.001" value={ins.cantidad} onChange={e => updateInsumo(recModal._id, iIdx, "cantidad", e.target.value)} placeholder="0" style={{ minWidth: 80 }} />
                            </td>
                            <td style={{ padding: "4px 10px", fontSize: 13, color: "#636366" }}>
                              {ins.unidad || skuInfo?.unidadVenta || "—"}
                            </td>
                            <td style={{ textAlign: "center", padding: "4px 6px" }}>
                              <button onClick={() => removeInsumo(recModal._id, iIdx)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30" }}><X size={14} /></button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <button className="btn-addrow" style={{ marginTop: 10 }} onClick={() => addInsumo(recModal._id)}>
                  <Plus size={14} /> Agregar insumo
                </button>
              </div>
              <div style={{ padding: "14px 22px", borderTop: "1px solid rgba(0,0,0,0.07)", display: "flex", alignItems: "center", gap: 8 }}>
                <button className="btn-ghost" style={{ fontSize: 13 }} onClick={descargarTemplateInsumos}><Download size={13} /> Plantilla</button>
                <label style={{ display: "inline-flex" }}>
                  <span className="btn-ghost" style={{ fontSize: 13, cursor: "pointer" }}><Upload size={13} /> Importar Excel</span>
                  <input ref={insumoImportRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => importarInsumos(recModal._id, e)} />
                </label>
                <div style={{ flex: 1 }} />
                <button className="btn-primary" style={{ fontSize: 14, padding: "9px 22px" }} onClick={() => setInsumoModal(null)}>Listo</button>
              </div>
            </div>
          </div>
        )}
      </main>
    </AppShell>
  );
}

function VistaCreacionCombo({ perfil, session, vista, setVista }) {
  const [combos, setCombos] = useState([emptySkuRow()]);
  const [expandedRow, setExpandedRow] = useState(null);
  const [comboModal, setComboModal] = useState(null);
  const [comboError, setComboError] = useState("");
  const [colGroup, setColGroup] = useState("basicos");
  const solicitante = perfil?.nombre || "";
  const [enviando, setEnviando] = useState(false);
  const [enviado, setEnviado] = useState(null);
  const [cat, setCat] = useState({ grupos: [], modelos: [], marcas: [], fabricantes: [] });
  const importRef = useRef();
  const comboImportRef = useRef();

  useEffect(() => {
    (async () => {
      const [{ data: grupos }, { data: modelos }, { data: marcas }, { data: fabricantes }] = await Promise.all([
        supabase.from("grupos_articulo").select("codigo,nombre"),
        supabase.from("materiales_modelo").select("codigo,nombre"),
        supabase.from("marcas").select("codigo,nombre"),
        supabase.from("fabricantes").select("codigo,nombre"),
      ]);
      setCat({ grupos: grupos || [], modelos: modelos || [], marcas: marcas || [], fabricantes: fabricantes || [] });
    })();
  }, []);

  const updateCombo  = (id, k, v) => setCombos(s => s.map(x => x._id === id ? { ...x, [k]: v } : x));
  const addCombo     = () => { if (combos.length >= 80) return; setCombos(s => [...s, emptySkuRow()]); };
  const removeCombo  = (id) => setCombos(s => s.filter(x => x._id !== id));

  const addFotos   = (id, tipo, files) => setCombos(s => s.map(x => x._id !== id ? x : { ...x, [`fotos_${tipo}`]: [...x[`fotos_${tipo}`], ...Array.from(files)] }));
  const removeFoto = (id, tipo, i)     => setCombos(s => s.map(x => x._id !== id ? x : { ...x, [`fotos_${tipo}`]: x[`fotos_${tipo}`].filter((_, j) => j !== i) }));

  const addPaso         = (comboId)               => setCombos(s => s.map(x => x._id !== comboId ? x : { ...x, pasos: [...x.pasos, emptyPaso()] }));
  const removePaso      = (comboId, pid)           => setCombos(s => s.map(x => x._id !== comboId ? x : { ...x, pasos: x.pasos.filter(p => p._id !== pid) }));
  const updatePaso      = (comboId, pid, k, v)     => setCombos(s => s.map(x => x._id !== comboId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, [k]: v }) }));
  const addProductoComboP   = (comboId, pid)           => setCombos(s => s.map(x => x._id !== comboId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, productos: [...p.productos, emptyProductoCombo()] }) }));
  const removeProductoCombo = (comboId, pid, prodId)   => setCombos(s => s.map(x => x._id !== comboId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, productos: p.productos.filter(pr => pr._id !== prodId) }) }));
  const updateProductoCombo = (comboId, pid, prodId, k, v) => setCombos(s => s.map(x => x._id !== comboId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, productos: p.productos.map(pr => pr._id !== prodId ? pr : { ...pr, [k]: v }) }) }));

  const descargarTemplateCombo = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Paso Nombre", "Obligatorio", "% Beneficio", "SKU", "Delta Precio"],
      ["Paso 1", "Si", 50, "132132", 0],
      ["Paso 1", "Si", 50, "133133", 200],
      ["Paso 2", "Si", 25, "135400", 1000],
    ]);
    ws["!cols"] = [{ wch: 14 }, { wch: 12 }, { wch: 13 }, { wch: 12 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pasos");
    XLSX.writeFile(wb, "template_pasos_combo.xlsx");
  };

  const importarCombo = (comboId, e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const pasoMap = {};
      const pasoOrder = [];
      raw.slice(1).filter(r => r[3]).forEach(r => {
        const nombre = String(r[0] ?? "").trim();
        const key = nombre || `Paso ${pasoOrder.length + 1}`;
        if (!pasoMap[key]) {
          const p = { ...emptyPaso(), nombre: key, obligatorio: String(r[1] ?? "Si").trim() === "No" ? "No" : "Si", pct_beneficio: String(r[2] ?? "").trim(), productos: [] };
          pasoMap[key] = p;
          pasoOrder.push(key);
        }
        pasoMap[key].productos.push({ _id: Math.random().toString(36).slice(2), sku: String(r[3]).trim(), delta_precio: String(r[4] ?? "").trim() });
      });
      const pasos = pasoOrder.map(k => pasoMap[k]);
      if (pasos.length) setCombos(s => s.map(x => x._id !== comboId ? x : { ...x, pasos }));
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const importarExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const nuevos = raw.slice(1).filter(r => r[0]).slice(0, 80).map(row => {
        const s = emptySkuRow();
        s.nombre          = String(row[0] ?? "").trim();
        s.descripcion     = String(row[1] ?? "").trim();
        s.grupo_articulo  = String(row[2] ?? "").trim();
        s.material_modelo = String(row[3] ?? "").trim();
        s.marca           = String(row[4] ?? "").trim();
        s.fabricante      = String(row[5] ?? "").trim();
        return s;
      });
      if (nuevos.length) setCombos(nuevos);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };
  const descargarTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Nombre", "Descripcion", "Grupo Articulo", "Material Modelo", "Marca", "Fabricante"],
      ["Combo Almuerzo", "Combo almuerzo con bebida", "CMB", "MAT001", "MARCA01", "FAB001"],
    ]);
    ws["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Combos");
    XLSX.writeFile(wb, "template_nuevos_combos.xlsx");
  };

  const validarTodo = () => {
    let ok = true;
    const updated = combos.map(combo => {
      const errs = [];
      if (!combo.nombre.trim())    errs.push("Nombre requerido");
      if (!combo.grupo_articulo)   errs.push("Grupo artículo requerido");
      if (!combo.material_modelo)  errs.push("Material modelo requerido");
      if (!combo.marca)            errs.push("Marca requerida");
      if (!combo.fabricante)       errs.push("Fabricante requerido");
      if (combo.pasos.length === 0) errs.push("Debe tener al menos un paso");
      const sumOblig = combo.pasos.filter(p => p.obligatorio === "Si").reduce((acc, p) => acc + (parseFloat(p.pct_beneficio) || 0), 0);
      if (combo.pasos.some(p => p.obligatorio === "Si") && Math.round(sumOblig) !== 100) errs.push(`Pasos obligatorios suman ${sumOblig}% (debe ser 100%)`);
      if (errs.length) ok = false;
      return { ...combo, errores: errs };
    });
    setCombos(updated);
    return ok;
  };

  const handleEnviar = async () => {
    if (!validarTodo() || solicitante.trim().length < 3) return;
    setEnviando(true);
    const folio = "SAP-" + new Date().getFullYear() + "-" + String(Math.floor(100000 + Math.random() * 899999));
    const filas = combos.map(({ _id, errores, fotos_planograma, fotos_navegacion, ...data }) => data);
    const { error } = await supabase.from("solicitudes").insert({
      folio,
      solicitante_id:     session?.user?.id,
      solicitante_nombre: solicitante,
      solicitante_email:  session?.user?.email,
      estado:    "Enviada",
      planillas: [{ id: "creacion_combo", nombre: "Creación de Combo", filas }],
      historial: [{ estado: "Enviada", fecha: new Date().toISOString() }],
    });
    setEnviando(false);
    if (!error) setEnviado({ folio, total: combos.length });
    else alert("Error al guardar: " + JSON.stringify(error));
  };

  if (enviado) return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil}>
      <main className="success-wrap">
        <div className="success-check"><CheckCircle2 size={64} strokeWidth={1.3} /></div>
        <h1 className="success-title">Solicitud enviada.</h1>
        <p className="success-sub">Folio <strong>{enviado.folio}</strong> · {enviado.total} combo{enviado.total > 1 ? "s" : ""} a crear.</p>
        <button className="btn-primary" style={{ marginTop: 24 }} onClick={() => { setEnviado(null); setCombos([emptySkuRow()]); }}>Nueva solicitud</button>
      </main>
    </AppShell>
  );

  const combosConError = combos.filter(c => c.errores.length > 0);

  const thSt = { padding: "8px 10px", textAlign: "left", fontSize: 11, fontWeight: 600, color: "#86868b", textTransform: "uppercase", letterSpacing: ".05em", borderBottom: "1px solid rgba(200,205,230,0.4)", background: "#f8f8fc", whiteSpace: "nowrap" };

  const COL_GROUPS = [
    { id: "basicos", label: "Datos básicos" },
    { id: "marcas",  label: "Marca / Prov." },
    { id: "fotos",   label: "Fotos / Flags" },
  ];

  const panelDerecho = (
    <div style={{ padding: "28px 20px", display: "flex", flexDirection: "column", gap: 16 }}>
      <div>
        <div style={{ fontSize: 13, fontWeight: 700, color: "#3a3a3c", marginBottom: 12 }}>Resumen</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span style={{ color: "#636366" }}>Combos</span>
            <strong>{combos.length} / 80</strong>
          </div>
          {combosConError.length > 0 && (
            <div style={{ fontSize: 12, color: "#c2271c", background: "rgba(255,59,48,0.07)", borderRadius: 8, padding: "8px 10px", marginTop: 4 }}>
              {combosConError.length} combo{combosConError.length > 1 ? "s" : ""} con errores
            </div>
          )}
        </div>
      </div>
      <div style={{ borderTop: "1px solid rgba(0,0,0,0.08)", paddingTop: 16 }}>
        <div style={{ fontSize: 12, color: "#636366", marginBottom: 6 }}>Solicitante</div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <User size={14} color="#636366" />
          <span style={{ fontSize: 13, color: solicitante ? "#3a3a3c" : "#aeaeb2", fontWeight: solicitante ? 500 : 400 }}>
            {solicitante || "Sin usuario autenticado"}
          </span>
        </div>
        <button className="btn-primary" style={{ width: "100%", marginBottom: 8 }} disabled={solicitante.trim().length < 3 || enviando} onClick={handleEnviar}>
          {enviando ? "Enviando…" : "Enviar solicitud"} <ArrowRight size={15} />
        </button>
      </div>
    </div>
  );

  return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil} rightPanel={panelDerecho}>
      <main className="sol-wrap">
        <div className="bloque">
          <div className="bloque-head">
            <span className="card-icon sm"><Gift size={18} strokeWidth={1.7} /></span>
            <div className="bloque-tit">
              <strong>Nuevo Combo</strong>
              <span>{combosConError.length > 0 ? `${combosConError.length} con error` : "Pendiente de validación"}</span>
            </div>
            <div className="bloque-acciones">
              <button className="btn-ghost" onClick={descargarTemplate}><Download size={14} /> Plantilla</button>
              <input type="file" accept=".xlsx,.xls" hidden ref={importRef} onChange={importarExcel} />
              <button className="btn-ghost" onClick={() => importRef.current?.click()}><Upload size={14} /> Importar Excel</button>
            </div>
          </div>

          <div style={{ display: "flex", gap: 6, padding: "8px 12px", borderBottom: "1px solid rgba(0,0,0,0.06)", flexWrap: "wrap" }}>
            {COL_GROUPS.map(g => (
              <button key={g.id} onClick={() => setColGroup(g.id)} style={{
                fontSize: 12, padding: "4px 14px", borderRadius: 980, border: "1px solid",
                borderColor: colGroup === g.id ? "#af52de" : "rgba(0,0,0,0.12)",
                background: colGroup === g.id ? "#af52de" : "transparent",
                color: colGroup === g.id ? "#fff" : "#515154",
                cursor: "pointer", fontWeight: colGroup === g.id ? 600 : 400, transition: "all .15s",
              }}>{g.label}</button>
            ))}
          </div>

          <div className="grilla-scroll">
            <table className="grilla">
              <thead>
                <tr>
                  <th className="th-n">#</th>
                  <th style={{ minWidth: 140 }}>ID Creación</th>
                  {colGroup === "basicos" && <>
                    <th style={{ minWidth: 175 }}>Nombre</th>
                    <th style={{ minWidth: 185 }}>Descripción</th>
                    <th style={{ minWidth: 110 }}>Pasos</th>
                  </>}
                  {colGroup === "marcas" && <>
                    <th style={{ minWidth: 140 }}>Grupo Art.</th>
                    <th style={{ minWidth: 140 }}>Mat. Modelo</th>
                    <th style={{ minWidth: 140 }}>Marca</th>
                    <th style={{ minWidth: 140 }}>Fabricante</th>
                  </>}
                  {colGroup === "fotos" && <>
                    <th style={{ minWidth: 65, textAlign: "center" }}>Plan.</th>
                    <th style={{ minWidth: 65, textAlign: "center" }}>Nav.</th>
                    <th style={{ minWidth: 110 }}>Fotos</th>
                  </>}
                  <th className="th-x"></th>
                </tr>
              </thead>
              <tbody>
                {combos.map((combo, idx) => {
                  const hayError = combo.errores.length > 0;
                  const expanded = expandedRow === combo._id;
                  const NCOLS_NOW = colGroup === "basicos" ? 6 : colGroup === "marcas" ? 7 : 6;
                  return (
                    <React.Fragment key={combo._id}>
                      <tr className={hayError ? "g-err" : ""}>
                        <td className="td-n">
                          {hayError ? <AlertTriangle size={13} className="ic-err" /> : <Pencil size={12} className="ic-dim" />}
                        </td>
                        <td>
                          <span style={{ fontFamily: "monospace", fontSize: 11, background: "#f0f0f5", color: "#5e5ce6", borderRadius: 6, padding: "2px 7px", whiteSpace: "nowrap" }}>
                            {combo.id_creacion}
                          </span>
                        </td>
                        {colGroup === "basicos" && <>
                          <td><input className="celda" value={combo.nombre} onChange={e => updateCombo(combo._id, "nombre", e.target.value)} placeholder="Nombre del combo" /></td>
                          <td><input className="celda" value={combo.descripcion} onChange={e => updateCombo(combo._id, "descripcion", e.target.value)} placeholder="Descripción" /></td>
                          <td>
                            <button className="dz-clear dup" style={{ padding: "4px 8px", fontSize: 12, whiteSpace: "nowrap", borderColor: "rgba(175,82,222,0.35)", color: "#af52de" }} onClick={() => { setComboError(""); setComboModal(combo._id); }}>
                              <Layers size={12} /> {combo.pasos.length > 0 ? `${combo.pasos.length} paso${combo.pasos.length > 1 ? "s" : ""}` : "Pasos"}
                            </button>
                          </td>
                        </>}
                        {colGroup === "marcas" && <>
                          <td>
                            <select className="celda" value={combo.grupo_articulo} onChange={e => updateCombo(combo._id, "grupo_articulo", e.target.value)}>
                              <option value="">—</option>
                              {cat.grupos.map(g => <option key={g.codigo} value={g.codigo}>{g.codigo} — {g.nombre}</option>)}
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={combo.material_modelo} onChange={e => updateCombo(combo._id, "material_modelo", e.target.value)}>
                              <option value="">—</option>
                              {cat.modelos.map(m => <option key={m.codigo} value={m.codigo}>{m.codigo} — {m.nombre}</option>)}
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={combo.marca} onChange={e => updateCombo(combo._id, "marca", e.target.value)}>
                              <option value="">—</option>
                              {cat.marcas.map(m => <option key={m.codigo} value={m.codigo}>{m.codigo} — {m.nombre}</option>)}
                              <option value="MARCA NO CREADA">NO CREADA</option>
                            </select>
                          </td>
                          <td>
                            <select className="celda" value={combo.fabricante} onChange={e => updateCombo(combo._id, "fabricante", e.target.value)}>
                              <option value="">—</option>
                              {cat.fabricantes.map(f => <option key={f.codigo} value={f.codigo}>{f.codigo} — {f.nombre}</option>)}
                              <option value="FABRICANTE NO CREADO">NO CREADO</option>
                            </select>
                          </td>
                        </>}
                        {colGroup === "fotos" && <>
                          <td style={{ textAlign: "center" }}>
                            <input type="checkbox" checked={combo.flag_planograma} onChange={e => updateCombo(combo._id, "flag_planograma", e.target.checked)} />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <input type="checkbox" checked={combo.flag_navegacion} onChange={e => updateCombo(combo._id, "flag_navegacion", e.target.checked)} />
                          </td>
                          <td>
                            <button className="dz-clear dup" style={{ padding: "4px 8px", fontSize: 12, whiteSpace: "nowrap" }} onClick={() => setExpandedRow(expanded ? null : combo._id)}>
                              <Upload size={12} /> {combo.fotos_planograma.length + combo.fotos_navegacion.length > 0 ? `${combo.fotos_planograma.length + combo.fotos_navegacion.length} foto${combo.fotos_planograma.length + combo.fotos_navegacion.length > 1 ? "s" : ""}` : "Adjuntar"}
                            </button>
                          </td>
                        </>}
                        <td className="td-x">
                          <span className="fila-btns">
                            {combos.length > 1 && (
                              <button className="dz-clear" onClick={() => removeCombo(combo._id)} title="Eliminar"><Trash2 size={13} /></button>
                            )}
                          </span>
                        </td>
                      </tr>
                      {hayError && (
                        <tr className="g-err-detalle">
                          <td></td>
                          <td colSpan={NCOLS_NOW - 1}>{combo.errores.join(" · ")}</td>
                        </tr>
                      )}
                      {expanded && colGroup === "fotos" && (
                        <tr>
                          <td></td>
                          <td colSpan={NCOLS_NOW - 1} style={{ padding: "14px 12px 18px", background: "#f8f8fc", borderBottom: "1px solid rgba(0,0,0,0.06)" }}>
                            <div style={{ display: "flex", gap: 14 }}>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 11, color: "#636366", marginBottom: 6 }}>Planograma · JPG (mín. 3)</div>
                                <label style={{ display: "block", border: "1.5px dashed rgba(0,0,0,0.15)", borderRadius: 10, padding: "10px", cursor: "pointer", textAlign: "center", fontSize: 12, color: "#636366" }}>
                                  <Upload size={14} style={{ display: "block", margin: "0 auto 4px" }} />
                                  {combo.fotos_planograma.length > 0 ? `${combo.fotos_planograma.length} foto${combo.fotos_planograma.length > 1 ? "s" : ""}` : "Seleccionar"}
                                  <input type="file" multiple accept="image/jpeg,image/jpg" style={{ display: "none" }} onChange={e => addFotos(combo._id, "planograma", e.target.files)} />
                                </label>
                                {combo.fotos_planograma.map((f, fi) => (
                                  <div key={fi} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#636366", marginTop: 3 }}>
                                    <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
                                    <button onClick={() => removeFoto(combo._id, "planograma", fi)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30", padding: 0 }}><X size={12} /></button>
                                  </div>
                                ))}
                              </div>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 11, color: "#636366", marginBottom: 6 }}>Navegación · JPG (mín. 3)</div>
                                <label style={{ display: "block", border: "1.5px dashed rgba(0,0,0,0.15)", borderRadius: 10, padding: "10px", cursor: "pointer", textAlign: "center", fontSize: 12, color: "#636366" }}>
                                  <Upload size={14} style={{ display: "block", margin: "0 auto 4px" }} />
                                  {combo.fotos_navegacion.length > 0 ? `${combo.fotos_navegacion.length} foto${combo.fotos_navegacion.length > 1 ? "s" : ""}` : "Seleccionar"}
                                  <input type="file" multiple accept="image/jpeg,image/jpg" style={{ display: "none" }} onChange={e => addFotos(combo._id, "navegacion", e.target.files)} />
                                </label>
                                {combo.fotos_navegacion.map((f, fi) => (
                                  <div key={fi} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#636366", marginTop: 3 }}>
                                    <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.name}</span>
                                    <button onClick={() => removeFoto(combo._id, "navegacion", fi)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30", padding: 0 }}><X size={12} /></button>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button className="btn-addrow" onClick={addCombo} disabled={combos.length >= 80}>
            <Plus size={14} /> Agregar combo
          </button>
        </div>

        {/* Modal pasos de combo */}
        {comboModal && (() => {
          const comboC = combos.find(c => c._id === comboModal);
          if (!comboC) return null;
          const sumOblig = comboC.pasos.filter(p => p.obligatorio === "Si").reduce((acc, p) => acc + (parseFloat(p.pct_beneficio) || 0), 0);
          const hayOblig = comboC.pasos.some(p => p.obligatorio === "Si");
          const pctOk    = !hayOblig || Math.round(sumOblig) === 100;
          return (
            <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}
                 onClick={e => { if (e.target === e.currentTarget) { setComboError(""); setComboModal(null); } }}>
              <div style={{ background: "#fff", borderRadius: 20, width: "100%", maxWidth: 800, maxHeight: "85vh", display: "flex", flexDirection: "column", overflow: "hidden", boxShadow: "0 20px 60px rgba(0,0,0,0.22)" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "18px 22px", borderBottom: "1px solid rgba(0,0,0,0.07)" }}>
                  <Layers size={18} color="#af52de" />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, fontSize: 15 }}>Pasos del combo</div>
                    <div style={{ fontSize: 12, color: "#636366" }}>
                      {comboC.nombre || "Combo sin nombre"} · <span style={{ fontFamily: "monospace", color: "#5e5ce6" }}>{comboC.id_creacion}</span>
                    </div>
                  </div>
                  {hayOblig && (
                    <span style={{ fontSize: 12, fontWeight: 600, padding: "3px 10px", borderRadius: 980, background: pctOk ? "rgba(52,199,89,0.12)" : "rgba(255,59,48,0.1)", color: pctOk ? "#34c759" : "#ff3b30" }}>
                      Oblig.: {sumOblig}% {pctOk ? "✓" : "≠ 100%"}
                    </span>
                  )}
                  <button onClick={() => { setComboError(""); setComboModal(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "#636366", padding: 4 }}><X size={20} /></button>
                </div>
                <div style={{ flex: 1, overflow: "auto", padding: "16px 22px" }}>
                  {comboC.pasos.length === 0 && (
                    <div style={{ textAlign: "center", color: "#aeaeb2", padding: "32px 0", fontSize: 14 }}>Sin pasos aún. Agrega el primero abajo.</div>
                  )}
                  {comboC.pasos.map((paso, pIdx) => (
                    <div key={paso._id} style={{ border: "1px solid rgba(0,0,0,0.09)", borderRadius: 12, overflow: "hidden", marginBottom: 12 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 14px", background: "#f8f8fc", borderBottom: "1px solid rgba(0,0,0,0.06)" }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: "#af52de", minWidth: 52 }}>PASO {pIdx + 1}</span>
                        <input
                          className="celda" value={paso.nombre}
                          onChange={e => updatePaso(comboC._id, paso._id, "nombre", e.target.value)}
                          placeholder="Nombre del paso (ej: Principal, Bebida…)"
                          style={{ flex: 1, fontSize: 13 }}
                        />
                        <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: "#3a3a3c", whiteSpace: "nowrap" }}>
                          Obligatorio:
                          <select className="celda" value={paso.obligatorio} onChange={e => updatePaso(comboC._id, paso._id, "obligatorio", e.target.value)} style={{ width: 60 }}>
                            <option>Si</option>
                            <option>No</option>
                          </select>
                        </label>
                        {paso.obligatorio === "Si" && (
                          <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: "#3a3a3c", whiteSpace: "nowrap" }}>
                            % Beneficio:
                            <input
                              className="celda" type="number" min="0" max="100" step="1"
                              value={paso.pct_beneficio}
                              onChange={e => updatePaso(comboC._id, paso._id, "pct_beneficio", e.target.value)}
                              placeholder="0"
                              style={{ width: 58, textAlign: "right" }}
                            />
                          </label>
                        )}
                        <button onClick={() => removePaso(comboC._id, paso._id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30", padding: "2px 4px", marginLeft: 4 }}><X size={15} /></button>
                      </div>
                      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
                        <thead>
                          <tr>
                            <th style={{ ...thSt, width: 120 }}>SKU</th>
                            <th style={thSt}>Nombre</th>
                            <th style={{ ...thSt, width: 110, textAlign: "right" }}>Delta Precio</th>
                            <th style={{ ...thSt, width: 36 }}></th>
                          </tr>
                        </thead>
                        <tbody>
                          {paso.productos.map(prod => {
                            const skuInfo = MAESTROS.skus?.[prod.sku];
                            return (
                              <tr key={prod._id} style={{ borderBottom: "1px solid rgba(0,0,0,0.04)" }}>
                                <td style={{ padding: "3px 6px" }}>
                                  <input className="celda" value={prod.sku}
                                    onChange={e => updateProductoCombo(comboC._id, paso._id, prod._id, "sku", e.target.value)}
                                    placeholder="Código SKU" style={{ minWidth: 100 }} />
                                </td>
                                <td style={{ padding: "4px 10px" }}>
                                  {skuInfo ? <span style={{ color: "#af52de" }}>{skuInfo.nombre}</span>
                                   : prod.sku ? <span style={{ color: "#ff3b30" }}>No encontrado</span>
                                   : <span style={{ color: "#c7c7cc" }}>—</span>}
                                </td>
                                <td style={{ padding: "3px 6px" }}>
                                  <input className="celda" type="number" step="0.01" value={prod.delta_precio}
                                    onChange={e => updateProductoCombo(comboC._id, paso._id, prod._id, "delta_precio", e.target.value)}
                                    placeholder="0" style={{ minWidth: 80, textAlign: "right" }} />
                                </td>
                                <td style={{ textAlign: "center", padding: "4px 6px" }}>
                                  <button onClick={() => removeProductoCombo(comboC._id, paso._id, prod._id)}
                                    style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30" }}><X size={14} /></button>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                      <button className="btn-addrow" style={{ margin: "6px 10px 10px" }} onClick={() => addProductoComboP(comboC._id, paso._id)}>
                        <Plus size={13} /> Agregar SKU
                      </button>
                    </div>
                  ))}
                  <button className="btn-addrow" onClick={() => addPaso(comboC._id)}>
                    <Plus size={14} /> Agregar paso
                  </button>
                </div>
                <div style={{ borderTop: "1px solid rgba(0,0,0,0.07)" }}>
                  {comboError && (
                    <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 22px", background: "rgba(255,59,48,0.07)", borderBottom: "1px solid rgba(255,59,48,0.12)" }}>
                      <AlertTriangle size={14} color="#ff3b30" />
                      <span style={{ fontSize: 13, color: "#ff3b30", fontWeight: 500 }}>{comboError}</span>
                    </div>
                  )}
                  <div style={{ padding: "14px 22px", display: "flex", alignItems: "center", gap: 8 }}>
                    <button className="btn-ghost" style={{ fontSize: 13 }} onClick={descargarTemplateCombo}><Download size={13} /> Plantilla</button>
                    <label style={{ display: "inline-flex" }}>
                      <span className="btn-ghost" style={{ fontSize: 13, cursor: "pointer" }}><Upload size={13} /> Importar Excel</span>
                      <input ref={comboImportRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => importarCombo(comboC._id, e)} />
                    </label>
                    <div style={{ flex: 1 }} />
                    <button className="btn-primary" style={{ fontSize: 14, padding: "9px 22px" }} onClick={() => {
                      if (hayOblig && !pctOk) {
                        setComboError(`Los pasos obligatorios suman ${sumOblig}% — deben sumar exactamente 100%.`);
                        return;
                      }
                      setComboError("");
                      setComboModal(null);
                    }}>Listo</button>
                  </div>
                </div>
              </div>
            </div>
          );
        })()}
      </main>
    </AppShell>
  );
}

const SRC_DESC = {
  "centro_nombre":      "centros → nombre",
  "sku_desc":           "SKUs → descripción",
  "sku_unidadVenta":    "SKUs → unidad de venta",
  "sku_unidadCondicion":"SKUs → unidad de condición",
  "centro_orgVenta":    "centros → org. de venta",
  "prov_nombre":        "proveedores → nombre",
  "status_signif":      "estado → significado",
};

function VistaMantenedorPlanillas() {
  const [abierta, setAbierta] = useState(null);

  return (
    <main style={{ maxWidth: 900, margin: "0 auto", padding: "32px 20px" }}>
      <h1 style={{ fontSize: 22, fontWeight: 700, marginBottom: 6 }}>Mantenedor de Planillas</h1>
      <p style={{ color: "#636366", fontSize: 14, marginBottom: 28 }}>
        Estructura de cada planilla: datos que ingresa el usuario, datos auto-completados desde bases maestras y columnas de salida para carga LSMW en SAP.
      </p>

      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {PLANILLAS.map(pl => {
          const open = abierta === pl.id;
          const Icon = pl.icon;
          return (
            <div key={pl.id} style={{ border: "1px solid rgba(0,0,0,0.1)", borderRadius: 14, overflow: "hidden", background: "#fff" }}>
              <button
                onClick={() => setAbierta(open ? null : pl.id)}
                style={{ width: "100%", display: "flex", alignItems: "center", gap: 12, padding: "16px 20px", background: "none", border: "none", cursor: "pointer", textAlign: "left" }}
              >
                <span style={{ width: 36, height: 36, borderRadius: 10, background: "rgba(91,141,238,0.1)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <Icon size={18} color="#5b8dee" />
                </span>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600, fontSize: 15 }}>{pl.nombre}</div>
                  <div style={{ fontSize: 13, color: "#636366" }}>{pl.desc}</div>
                </div>
                <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                  <span style={{ fontSize: 12, padding: "3px 8px", borderRadius: 20, background: "rgba(91,141,238,0.1)", color: "#5b8dee" }}>{pl.userCols.length} usuario</span>
                  <span style={{ fontSize: 12, padding: "3px 8px", borderRadius: 20, background: "rgba(52,199,89,0.1)", color: "#248a3d" }}>{pl.autoCols.length} auto</span>
                  <span style={{ fontSize: 12, padding: "3px 8px", borderRadius: 20, background: "rgba(255,149,0,0.12)", color: "#b25000" }}>{pl.salidaLabels.length} LSMW</span>
                </div>
                {open ? <ChevronDown size={16} color="#636366" /> : <ChevronRight size={16} color="#636366" />}
              </button>

              {open && (
                <div style={{ padding: "0 20px 20px", borderTop: "1px solid rgba(0,0,0,0.06)" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginTop: 16 }}>

                    {/* Columnas del usuario */}
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#5b8dee", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>
                        Ingresa el usuario
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        {pl.userCols.map(c => (
                          <div key={c.key} style={{ background: "rgba(91,141,238,0.07)", border: "1px solid rgba(91,141,238,0.2)", borderRadius: 8, padding: "8px 12px" }}>
                            <div style={{ fontWeight: 600, fontSize: 13 }}>{c.label}</div>
                            <div style={{ fontSize: 12, color: "#636366" }}>
                              {c.centro ? "Código de centro / cluster" : c.sku ? "Código SKU" : c.proveedor ? "Código proveedor" : c.fecha ? "Fecha DD-MM-AAAA" : c.num ? "Numérico" : "Texto"}
                              {c.def ? <span style={{ marginLeft: 6, color: "#5b8dee" }}>def: {c.def}</span> : null}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Columnas auto */}
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#248a3d", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>
                        Auto-completado
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        {pl.autoCols.map(c => (
                          <div key={c.key} style={{ background: "rgba(52,199,89,0.07)", border: "1px solid rgba(52,199,89,0.2)", borderRadius: 8, padding: "8px 12px" }}>
                            <div style={{ fontWeight: 600, fontSize: 13 }}>{c.label}</div>
                            <div style={{ fontSize: 12, color: "#636366" }}>{SRC_DESC[c.src] || c.src}</div>
                          </div>
                        ))}
                        {(pl.agentCols || []).map(c => (
                          <div key={c.key} style={{ background: "rgba(175,82,222,0.07)", border: "1px solid rgba(175,82,222,0.2)", borderRadius: 8, padding: "8px 12px" }}>
                            <div style={{ fontWeight: 600, fontSize: 13 }}>{c.label}</div>
                            <div style={{ fontSize: 12, color: "#636366" }}>Valor fijo: <strong>{c.def}</strong></div>
                          </div>
                        ))}
                        {pl.autoCols.length === 0 && (pl.agentCols || []).length === 0 && (
                          <div style={{ fontSize: 13, color: "#aeaeb2", fontStyle: "italic" }}>Sin columnas auto</div>
                        )}
                      </div>
                    </div>

                    {/* Columnas LSMW */}
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 600, color: "#b25000", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>
                        Salida LSMW (SAP)
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        {pl.salidaLabels.map((lbl, i) => (
                          <div key={i} style={{ background: "rgba(255,149,0,0.07)", border: "1px solid rgba(255,149,0,0.2)", borderRadius: 8, padding: "8px 12px" }}>
                            <div style={{ fontWeight: 600, fontSize: 13 }}>{lbl}</div>
                            <div style={{ fontSize: 12, color: "#aeaeb2", fontFamily: "monospace" }}>{pl.salida[i]}</div>
                          </div>
                        ))}
                      </div>
                    </div>

                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </main>
  );
}

/* ─── helpers modificación ─── */
const emptyModReceta = () => ({ _id: Math.random().toString(36).slice(2), codigo_sap: "", nombre: "", insumos: [], errores: [] });
const emptyModCombo  = () => ({ _id: Math.random().toString(36).slice(2), codigo_sap: "", nombre: "", pasos: [],   errores: [] });

function VistaModReceta({ perfil, session, vista, setVista }) {
  const [filas, setFilas]       = useState([emptyModReceta()]);
  const [insumoModal, setInsumoModal] = useState(null);
  const solicitante             = perfil?.nombre || "";
  const [enviando, setEnviando] = useState(false);
  const [enviado, setEnviado]   = useState(null);
  const insumoImportRef         = useRef();

  const updateFila   = (id, k, v) => setFilas(s => s.map(x => x._id === id ? { ...x, [k]: v } : x));
  const addFila      = () => { if (filas.length >= 80) return; setFilas(s => [...s, emptyModReceta()]); };
  const removeFila   = (id) => setFilas(s => s.filter(x => x._id !== id));

  const addInsumo    = (id)          => setFilas(s => s.map(x => x._id !== id ? x : { ...x, insumos: [...x.insumos, { sku: "", cantidad: "", unidad: "" }] }));
  const removeInsumo = (id, i)       => setFilas(s => s.map(x => x._id !== id ? x : { ...x, insumos: x.insumos.filter((_, j) => j !== i) }));
  const updateInsumo = (id, i, k, v) => setFilas(s => s.map(x => {
    if (x._id !== id) return x;
    const ins = x.insumos.map((r, j) => {
      if (j !== i) return r;
      const u = { ...r, [k]: v };
      if (k === "sku") u.unidad = MAESTROS.skus?.[v]?.unidadVenta || "";
      return u;
    });
    return { ...x, insumos: ins };
  }));

  const descargarTemplateInsumos = () => {
    const ws = XLSX.utils.aoa_to_sheet([["Codigo SKU", "Cantidad"], ["10047", "2.5"]]);
    ws["!cols"] = [{ wch: 16 }, { wch: 12 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Insumos");
    XLSX.writeFile(wb, "template_insumos_receta.xlsx");
  };
  const importarInsumos = (id, e) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });
      const nuevos = raw.slice(1).filter(r => r[0]).map(r => {
        const skuCod = String(r[0]).trim();
        return { sku: skuCod, cantidad: String(r[1] ?? "").trim(), unidad: MAESTROS.skus?.[skuCod]?.unidadVenta || "" };
      });
      if (nuevos.length) setFilas(s => s.map(x => x._id !== id ? x : { ...x, insumos: nuevos }));
    };
    reader.readAsArrayBuffer(file); e.target.value = "";
  };

  const validarTodo = () => {
    let ok = true;
    const updated = filas.map(f => {
      const errs = [];
      if (!f.codigo_sap.trim()) errs.push("Código SAP requerido");
      if (!f.nombre.trim())     errs.push("Nombre requerido");
      if (f.insumos.length === 0) errs.push("Debe tener al menos un insumo");
      if (errs.length) ok = false;
      return { ...f, errores: errs };
    });
    setFilas(updated); return ok;
  };

  const handleEnviar = async () => {
    if (!validarTodo() || solicitante.trim().length < 3) return;
    setEnviando(true);
    const folio = "SAP-" + new Date().getFullYear() + "-" + String(Math.floor(100000 + Math.random() * 899999));
    const { error } = await supabase.from("solicitudes").insert({
      folio, solicitante_id: session?.user?.id, solicitante_nombre: solicitante,
      solicitante_email: session?.user?.email, estado: "Enviada",
      planillas: [{ id: "modificacion_receta", nombre: "Modificación de Receta", filas: filas.map(({ _id, errores, ...d }) => d) }],
      historial: [{ estado: "Enviada", fecha: new Date().toISOString() }],
    });
    setEnviando(false);
    if (!error) setEnviado({ folio, total: filas.length });
    else alert("Error al guardar: " + JSON.stringify(error));
  };

  if (enviado) return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil}>
      <main className="success-wrap">
        <div className="success-check"><CheckCircle2 size={64} strokeWidth={1.3} /></div>
        <h1 className="success-title">Solicitud enviada.</h1>
        <p className="success-sub">Folio <strong>{enviado.folio}</strong> · {enviado.total} receta{enviado.total > 1 ? "s" : ""} a modificar.</p>
        <button className="btn-primary" style={{ marginTop: 24 }} onClick={() => { setEnviado(null); setFilas([emptyModReceta()]); }}>Nueva solicitud</button>
      </main>
    </AppShell>
  );

  const conError  = filas.filter(f => f.errores.length > 0);
  const recModal  = insumoModal ? filas.find(f => f._id === insumoModal) : null;
  const thSt      = { padding: "8px 10px", textAlign: "left", fontSize: 11, fontWeight: 600, color: "#86868b", textTransform: "uppercase", letterSpacing: ".05em", borderBottom: "1px solid rgba(200,205,230,0.4)", background: "#f8f8fc", whiteSpace: "nowrap" };

  const panelDerecho = (
    <div style={{ padding: "28px 20px", display: "flex", flexDirection: "column", gap: 16 }}>
      <div>
        <div style={{ fontSize: 13, fontWeight: 700, color: "#3a3a3c", marginBottom: 12 }}>Resumen</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span style={{ color: "#636366" }}>Recetas</span><strong>{filas.length}</strong>
          </div>
          {conError.length > 0 && (
            <div style={{ fontSize: 12, color: "#c2271c", background: "rgba(255,59,48,0.07)", borderRadius: 8, padding: "8px 10px", marginTop: 4 }}>
              {conError.length} con errores
            </div>
          )}
        </div>
      </div>
      <div style={{ borderTop: "1px solid rgba(0,0,0,0.08)", paddingTop: 16 }}>
        <div style={{ fontSize: 12, color: "#636366", marginBottom: 6 }}>Solicitante</div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <User size={14} color="#636366" />
          <span style={{ fontSize: 13, color: solicitante ? "#3a3a3c" : "#aeaeb2", fontWeight: solicitante ? 500 : 400 }}>
            {solicitante || "Sin usuario autenticado"}
          </span>
        </div>
        <button className="btn-primary" style={{ width: "100%", marginBottom: 8 }} disabled={solicitante.trim().length < 3 || enviando} onClick={handleEnviar}>
          {enviando ? "Enviando…" : "Enviar solicitud"} <ArrowRight size={15} />
        </button>
      </div>
    </div>
  );

  return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil} rightPanel={panelDerecho}>
      <main className="sol-wrap">
        <div className="bloque">
          <div className="bloque-head">
            <span className="card-icon sm"><Pencil size={18} strokeWidth={1.7} /></span>
            <div className="bloque-tit">
              <strong>Modificar Receta</strong>
              <span>{conError.length > 0 ? `${conError.length} con error` : "Ingresa el código SAP y la nueva estructura de insumos"}</span>
            </div>
          </div>

          <div className="grilla-scroll">
            <table className="grilla">
              <thead>
                <tr>
                  <th className="th-n">#</th>
                  <th style={{ minWidth: 120 }}>Código SAP</th>
                  <th style={{ minWidth: 175 }}>Nombre receta</th>
                  <th style={{ minWidth: 110 }}>Insumos</th>
                  <th className="th-x"></th>
                </tr>
              </thead>
              <tbody>
                {filas.map((fila, idx) => {
                  const hayError = fila.errores.length > 0;
                  return (
                    <React.Fragment key={fila._id}>
                      <tr className={hayError ? "g-err" : ""}>
                        <td className="td-n">{hayError ? <AlertTriangle size={13} className="ic-err" /> : <Pencil size={12} className="ic-dim" />}</td>
                        <td><input className="celda" value={fila.codigo_sap} onChange={e => updateFila(fila._id, "codigo_sap", e.target.value)} placeholder="Ej: 132132" style={{ fontFamily: "monospace" }} /></td>
                        <td><input className="celda" value={fila.nombre} onChange={e => updateFila(fila._id, "nombre", e.target.value)} placeholder="Nombre de la receta" /></td>
                        <td>
                          <button className="dz-clear dup" style={{ padding: "4px 8px", fontSize: 12, whiteSpace: "nowrap" }} onClick={() => setInsumoModal(fila._id)}>
                            <Layers size={12} /> {fila.insumos.length > 0 ? `${fila.insumos.length} insumo${fila.insumos.length > 1 ? "s" : ""}` : "Insumos"}
                          </button>
                        </td>
                        <td className="td-x">
                          {filas.length > 1 && <button className="dz-clear" onClick={() => removeFila(fila._id)} title="Eliminar"><Trash2 size={13} /></button>}
                        </td>
                      </tr>
                      {hayError && (
                        <tr className="g-err-detalle">
                          <td></td>
                          <td colSpan={4}>{fila.errores.join(" · ")}</td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button className="btn-addrow" onClick={addFila} disabled={filas.length >= 80}>
            <Plus size={14} /> Agregar receta
          </button>
        </div>

        {/* Modal insumos */}
        {recModal && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}
               onClick={e => { if (e.target === e.currentTarget) setInsumoModal(null); }}>
            <div style={{ background: "#fff", borderRadius: 20, width: "100%", maxWidth: 660, maxHeight: "80vh", display: "flex", flexDirection: "column", overflow: "hidden", boxShadow: "0 20px 60px rgba(0,0,0,0.22)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "18px 22px", borderBottom: "1px solid rgba(0,0,0,0.07)" }}>
                <Layers size={18} color="#34c759" />
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>Nueva estructura de insumos</div>
                  <div style={{ fontSize: 12, color: "#636366" }}>
                    {recModal.nombre || "Receta sin nombre"} · <span style={{ fontFamily: "monospace", color: "#5e5ce6" }}>SAP {recModal.codigo_sap || "—"}</span>
                  </div>
                </div>
                <button onClick={() => setInsumoModal(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "#636366", padding: 4 }}><X size={20} /></button>
              </div>
              <div style={{ flex: 1, overflow: "auto", padding: "16px 22px" }}>
                <div style={{ fontSize: 12, color: "#636366", background: "rgba(91,141,238,0.07)", borderRadius: 10, padding: "10px 14px", marginBottom: 12 }}>
                  Define la nueva lista completa de insumos tal como debe quedar en SAP.
                </div>
                <div style={{ border: "1px solid rgba(0,0,0,0.07)", borderRadius: 12, overflow: "hidden" }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 14 }}>
                    <thead>
                      <tr>
                        <th style={thSt}>Código SKU</th>
                        <th style={thSt}>Nombre</th>
                        <th style={{ ...thSt, width: 110 }}>Cantidad</th>
                        <th style={{ ...thSt, width: 80 }}>Unidad</th>
                        <th style={{ ...thSt, width: 36 }}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {recModal.insumos.length === 0 && (
                        <tr><td colSpan={5} style={{ textAlign: "center", color: "#aeaeb2", padding: "22px 10px", fontSize: 13 }}>Sin insumos. Agrega el primero abajo.</td></tr>
                      )}
                      {recModal.insumos.map((ins, iIdx) => {
                        const skuInfo = MAESTROS.skus?.[ins.sku];
                        return (
                          <tr key={iIdx} style={{ borderBottom: "1px solid rgba(0,0,0,0.04)" }}>
                            <td style={{ padding: "3px 6px" }}><input className="celda" value={ins.sku} onChange={e => updateInsumo(recModal._id, iIdx, "sku", e.target.value)} placeholder="Código" style={{ minWidth: 100 }} /></td>
                            <td style={{ padding: "4px 10px" }}>
                              {skuInfo ? <span style={{ fontSize: 13, color: "#34c759" }}>{skuInfo.nombre}</span>
                               : ins.sku ? <span style={{ fontSize: 13, color: "#ff3b30" }}>No encontrado</span>
                               : <span style={{ fontSize: 13, color: "#c7c7cc" }}>—</span>}
                            </td>
                            <td style={{ padding: "3px 6px" }}><input className="celda" type="number" min="0" step="0.001" value={ins.cantidad} onChange={e => updateInsumo(recModal._id, iIdx, "cantidad", e.target.value)} placeholder="0" style={{ minWidth: 80 }} /></td>
                            <td style={{ padding: "4px 10px", fontSize: 13, color: "#636366" }}>{ins.unidad || skuInfo?.unidadVenta || "—"}</td>
                            <td style={{ textAlign: "center", padding: "4px 6px" }}><button onClick={() => removeInsumo(recModal._id, iIdx)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30" }}><X size={14} /></button></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <button className="btn-addrow" style={{ marginTop: 10 }} onClick={() => addInsumo(recModal._id)}><Plus size={14} /> Agregar insumo</button>
              </div>
              <div style={{ padding: "14px 22px", borderTop: "1px solid rgba(0,0,0,0.07)", display: "flex", alignItems: "center", gap: 8 }}>
                <button className="btn-ghost" style={{ fontSize: 13 }} onClick={descargarTemplateInsumos}><Download size={13} /> Plantilla</button>
                <label style={{ display: "inline-flex" }}>
                  <span className="btn-ghost" style={{ fontSize: 13, cursor: "pointer" }}><Upload size={13} /> Importar Excel</span>
                  <input ref={insumoImportRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => importarInsumos(recModal._id, e)} />
                </label>
                <div style={{ flex: 1 }} />
                <button className="btn-primary" style={{ fontSize: 14, padding: "9px 22px" }} onClick={() => setInsumoModal(null)}>Listo</button>
              </div>
            </div>
          </div>
        )}
      </main>
    </AppShell>
  );
}

function VistaModCombo({ perfil, session, vista, setVista }) {
  const [filas, setFilas]         = useState([emptyModCombo()]);
  const [comboModal, setComboModal] = useState(null);
  const [comboError, setComboError] = useState("");
  const solicitante               = perfil?.nombre || "";
  const [enviando, setEnviando]   = useState(false);
  const [enviado, setEnviado]     = useState(null);
  const comboImportRef            = useRef();

  const updateFila = (id, k, v) => setFilas(s => s.map(x => x._id === id ? { ...x, [k]: v } : x));
  const addFila    = () => { if (filas.length >= 80) return; setFilas(s => [...s, emptyModCombo()]); };
  const removeFila = (id) => setFilas(s => s.filter(x => x._id !== id));

  const addPaso         = (skuId)              => setFilas(s => s.map(x => x._id !== skuId ? x : { ...x, pasos: [...x.pasos, emptyPaso()] }));
  const removePaso      = (skuId, pid)         => setFilas(s => s.map(x => x._id !== skuId ? x : { ...x, pasos: x.pasos.filter(p => p._id !== pid) }));
  const updatePaso      = (skuId, pid, k, v)   => setFilas(s => s.map(x => x._id !== skuId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, [k]: v }) }));
  const addProd         = (skuId, pid)         => setFilas(s => s.map(x => x._id !== skuId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, productos: [...p.productos, emptyProductoCombo()] }) }));
  const removeProd      = (skuId, pid, prodId) => setFilas(s => s.map(x => x._id !== skuId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, productos: p.productos.filter(pr => pr._id !== prodId) }) }));
  const updateProd      = (skuId, pid, prodId, k, v) => setFilas(s => s.map(x => x._id !== skuId ? x : { ...x, pasos: x.pasos.map(p => p._id !== pid ? p : { ...p, productos: p.productos.map(pr => pr._id !== prodId ? pr : { ...pr, [k]: v }) }) }));

  const descargarTemplateCombo = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["Paso Nombre", "Obligatorio", "% Beneficio", "SKU", "Delta Precio"],
      ["Paso 1", "Si", 50, "132132", 0], ["Paso 1", "Si", 50, "133133", 200], ["Paso 2", "Si", 50, "135400", 0],
    ]);
    ws["!cols"] = [{ wch: 14 }, { wch: 12 }, { wch: 13 }, { wch: 12 }, { wch: 12 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Pasos");
    XLSX.writeFile(wb, "template_pasos_combo.xlsx");
  };
  const importarCombo = (skuId, e) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array" });
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });
      const pasoMap = {}; const pasoOrder = [];
      raw.slice(1).filter(r => r[3]).forEach(r => {
        const nombre = String(r[0] ?? "").trim();
        const key = nombre || `Paso ${pasoOrder.length + 1}`;
        if (!pasoMap[key]) { pasoMap[key] = { ...emptyPaso(), nombre: key, obligatorio: String(r[1] ?? "Si").trim() === "No" ? "No" : "Si", pct_beneficio: String(r[2] ?? "").trim(), productos: [] }; pasoOrder.push(key); }
        pasoMap[key].productos.push({ _id: Math.random().toString(36).slice(2), sku: String(r[3]).trim(), delta_precio: String(r[4] ?? "").trim() });
      });
      const pasos = pasoOrder.map(k => pasoMap[k]);
      if (pasos.length) setFilas(s => s.map(x => x._id !== skuId ? x : { ...x, pasos }));
    };
    reader.readAsArrayBuffer(file); e.target.value = "";
  };

  const validarTodo = () => {
    let ok = true;
    const updated = filas.map(f => {
      const errs = [];
      if (!f.codigo_sap.trim()) errs.push("Código SAP requerido");
      if (!f.nombre.trim())     errs.push("Nombre requerido");
      if (f.pasos.length === 0) errs.push("Debe tener al menos un paso");
      const sumOblig = f.pasos.filter(p => p.obligatorio === "Si").reduce((a, p) => a + (parseFloat(p.pct_beneficio) || 0), 0);
      if (f.pasos.some(p => p.obligatorio === "Si") && Math.round(sumOblig) !== 100) errs.push(`Pasos obligatorios suman ${sumOblig}% (debe ser 100%)`);
      if (errs.length) ok = false;
      return { ...f, errores: errs };
    });
    setFilas(updated); return ok;
  };

  const handleEnviar = async () => {
    if (!validarTodo() || solicitante.trim().length < 3) return;
    setEnviando(true);
    const folio = "SAP-" + new Date().getFullYear() + "-" + String(Math.floor(100000 + Math.random() * 899999));
    const { error } = await supabase.from("solicitudes").insert({
      folio, solicitante_id: session?.user?.id, solicitante_nombre: solicitante,
      solicitante_email: session?.user?.email, estado: "Enviada",
      planillas: [{ id: "modificacion_combo", nombre: "Modificación de Combo", filas: filas.map(({ _id, errores, ...d }) => d) }],
      historial: [{ estado: "Enviada", fecha: new Date().toISOString() }],
    });
    setEnviando(false);
    if (!error) setEnviado({ folio, total: filas.length });
    else alert("Error al guardar: " + JSON.stringify(error));
  };

  if (enviado) return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil}>
      <main className="success-wrap">
        <div className="success-check"><CheckCircle2 size={64} strokeWidth={1.3} /></div>
        <h1 className="success-title">Solicitud enviada.</h1>
        <p className="success-sub">Folio <strong>{enviado.folio}</strong> · {enviado.total} combo{enviado.total > 1 ? "s" : ""} a modificar.</p>
        <button className="btn-primary" style={{ marginTop: 24 }} onClick={() => { setEnviado(null); setFilas([emptyModCombo()]); }}>Nueva solicitud</button>
      </main>
    </AppShell>
  );

  const conError = filas.filter(f => f.errores.length > 0);
  const thSt     = { padding: "8px 10px", textAlign: "left", fontSize: 11, fontWeight: 600, color: "#86868b", textTransform: "uppercase", letterSpacing: ".05em", borderBottom: "1px solid rgba(200,205,230,0.4)", background: "#f8f8fc", whiteSpace: "nowrap" };

  const panelDerecho = (
    <div style={{ padding: "28px 20px", display: "flex", flexDirection: "column", gap: 16 }}>
      <div>
        <div style={{ fontSize: 13, fontWeight: 700, color: "#3a3a3c", marginBottom: 12 }}>Resumen</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span style={{ color: "#636366" }}>Combos</span><strong>{filas.length}</strong>
          </div>
          {conError.length > 0 && (
            <div style={{ fontSize: 12, color: "#c2271c", background: "rgba(255,59,48,0.07)", borderRadius: 8, padding: "8px 10px", marginTop: 4 }}>
              {conError.length} con errores
            </div>
          )}
        </div>
      </div>
      <div style={{ borderTop: "1px solid rgba(0,0,0,0.08)", paddingTop: 16 }}>
        <div style={{ fontSize: 12, color: "#636366", marginBottom: 6 }}>Solicitante</div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <User size={14} color="#636366" />
          <span style={{ fontSize: 13, color: solicitante ? "#3a3a3c" : "#aeaeb2", fontWeight: solicitante ? 500 : 400 }}>
            {solicitante || "Sin usuario autenticado"}
          </span>
        </div>
        <button className="btn-primary" style={{ width: "100%", marginBottom: 8 }} disabled={solicitante.trim().length < 3 || enviando} onClick={handleEnviar}>
          {enviando ? "Enviando…" : "Enviar solicitud"} <ArrowRight size={15} />
        </button>
      </div>
    </div>
  );

  return (
    <AppShell vista={vista} setVista={setVista} perfil={perfil} rightPanel={panelDerecho}>
      <main className="sol-wrap">
        <div className="bloque">
          <div className="bloque-head">
            <span className="card-icon sm"><Pencil size={18} strokeWidth={1.7} /></span>
            <div className="bloque-tit">
              <strong>Modificar Combo</strong>
              <span>{conError.length > 0 ? `${conError.length} con error` : "Ingresa el código SAP y la nueva estructura de pasos"}</span>
            </div>
          </div>

          <div className="grilla-scroll">
            <table className="grilla">
              <thead>
                <tr>
                  <th className="th-n">#</th>
                  <th style={{ minWidth: 120 }}>Código SAP</th>
                  <th style={{ minWidth: 175 }}>Nombre combo</th>
                  <th style={{ minWidth: 110 }}>Pasos</th>
                  <th className="th-x"></th>
                </tr>
              </thead>
              <tbody>
                {filas.map((fila) => {
                  const hayError = fila.errores.length > 0;
                  return (
                    <React.Fragment key={fila._id}>
                      <tr className={hayError ? "g-err" : ""}>
                        <td className="td-n">{hayError ? <AlertTriangle size={13} className="ic-err" /> : <Pencil size={12} className="ic-dim" />}</td>
                        <td><input className="celda" value={fila.codigo_sap} onChange={e => updateFila(fila._id, "codigo_sap", e.target.value)} placeholder="Ej: 132132" style={{ fontFamily: "monospace" }} /></td>
                        <td><input className="celda" value={fila.nombre} onChange={e => updateFila(fila._id, "nombre", e.target.value)} placeholder="Nombre del combo" /></td>
                        <td>
                          <button className="dz-clear dup" style={{ padding: "4px 8px", fontSize: 12, whiteSpace: "nowrap", borderColor: "rgba(175,82,222,0.35)", color: "#af52de" }}
                            onClick={() => { setComboError(""); setComboModal(fila._id); }}>
                            <Layers size={12} /> {fila.pasos.length > 0 ? `${fila.pasos.length} paso${fila.pasos.length > 1 ? "s" : ""}` : "Pasos"}
                          </button>
                        </td>
                        <td className="td-x">
                          {filas.length > 1 && <button className="dz-clear" onClick={() => removeFila(fila._id)} title="Eliminar"><Trash2 size={13} /></button>}
                        </td>
                      </tr>
                      {hayError && (
                        <tr className="g-err-detalle">
                          <td></td>
                          <td colSpan={4}>{fila.errores.join(" · ")}</td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button className="btn-addrow" onClick={addFila} disabled={filas.length >= 80}>
            <Plus size={14} /> Agregar combo
          </button>
        </div>

        {/* Modal pasos combo */}
        {comboModal && (() => {
          const filaC = filas.find(f => f._id === comboModal);
          if (!filaC) return null;
          const sumOblig = filaC.pasos.filter(p => p.obligatorio === "Si").reduce((a, p) => a + (parseFloat(p.pct_beneficio) || 0), 0);
          const hayOblig = filaC.pasos.some(p => p.obligatorio === "Si");
          const pctOk    = !hayOblig || Math.round(sumOblig) === 100;
          return (
            <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}
                 onClick={e => { if (e.target === e.currentTarget) { setComboError(""); setComboModal(null); } }}>
              <div style={{ background: "#fff", borderRadius: 20, width: "100%", maxWidth: 800, maxHeight: "85vh", display: "flex", flexDirection: "column", overflow: "hidden", boxShadow: "0 20px 60px rgba(0,0,0,0.22)" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "18px 22px", borderBottom: "1px solid rgba(0,0,0,0.07)" }}>
                  <Layers size={18} color="#af52de" />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, fontSize: 15 }}>Nueva estructura de pasos</div>
                    <div style={{ fontSize: 12, color: "#636366" }}>
                      {filaC.nombre || "Combo sin nombre"} · <span style={{ fontFamily: "monospace", color: "#5e5ce6" }}>SAP {filaC.codigo_sap || "—"}</span>
                    </div>
                  </div>
                  {hayOblig && (
                    <span style={{ fontSize: 12, fontWeight: 600, padding: "3px 10px", borderRadius: 980, background: pctOk ? "rgba(52,199,89,0.12)" : "rgba(255,59,48,0.1)", color: pctOk ? "#34c759" : "#ff3b30" }}>
                      Oblig.: {sumOblig}% {pctOk ? "✓" : "≠ 100%"}
                    </span>
                  )}
                  <button onClick={() => { setComboError(""); setComboModal(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "#636366", padding: 4 }}><X size={20} /></button>
                </div>
                <div style={{ flex: 1, overflow: "auto", padding: "16px 22px" }}>
                  <div style={{ fontSize: 12, color: "#636366", background: "rgba(175,82,222,0.07)", borderRadius: 10, padding: "10px 14px", marginBottom: 12 }}>
                    Define la nueva estructura completa de pasos y productos tal como debe quedar en SAP.
                  </div>
                  {filaC.pasos.length === 0 && (
                    <div style={{ textAlign: "center", color: "#aeaeb2", padding: "32px 0", fontSize: 14 }}>Sin pasos. Agrega el primero abajo.</div>
                  )}
                  {filaC.pasos.map((paso, pIdx) => (
                    <div key={paso._id} style={{ border: "1px solid rgba(0,0,0,0.09)", borderRadius: 12, overflow: "hidden", marginBottom: 12 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 14px", background: "#f8f8fc", borderBottom: "1px solid rgba(0,0,0,0.06)" }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: "#af52de", minWidth: 52 }}>PASO {pIdx + 1}</span>
                        <input className="celda" value={paso.nombre} onChange={e => updatePaso(filaC._id, paso._id, "nombre", e.target.value)} placeholder="Nombre del paso" style={{ flex: 1, fontSize: 13 }} />
                        <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: "#3a3a3c", whiteSpace: "nowrap" }}>
                          Obligatorio:
                          <select className="celda" value={paso.obligatorio} onChange={e => updatePaso(filaC._id, paso._id, "obligatorio", e.target.value)} style={{ width: 60 }}>
                            <option>Si</option><option>No</option>
                          </select>
                        </label>
                        {paso.obligatorio === "Si" && (
                          <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: "#3a3a3c", whiteSpace: "nowrap" }}>
                            % Beneficio:
                            <input className="celda" type="number" min="0" max="100" step="1" value={paso.pct_beneficio} onChange={e => updatePaso(filaC._id, paso._id, "pct_beneficio", e.target.value)} placeholder="0" style={{ width: 58, textAlign: "right" }} />
                          </label>
                        )}
                        <button onClick={() => removePaso(filaC._id, paso._id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30", padding: "2px 4px", marginLeft: 4 }}><X size={15} /></button>
                      </div>
                      <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 13 }}>
                        <thead>
                          <tr>
                            <th style={{ ...thSt, width: 120 }}>SKU</th>
                            <th style={thSt}>Nombre</th>
                            <th style={{ ...thSt, width: 110, textAlign: "right" }}>Delta Precio</th>
                            <th style={{ ...thSt, width: 36 }}></th>
                          </tr>
                        </thead>
                        <tbody>
                          {paso.productos.map(prod => {
                            const skuInfo = MAESTROS.skus?.[prod.sku];
                            return (
                              <tr key={prod._id} style={{ borderBottom: "1px solid rgba(0,0,0,0.04)" }}>
                                <td style={{ padding: "3px 6px" }}><input className="celda" value={prod.sku} onChange={e => updateProd(filaC._id, paso._id, prod._id, "sku", e.target.value)} placeholder="Código SKU" style={{ minWidth: 100 }} /></td>
                                <td style={{ padding: "4px 10px" }}>
                                  {skuInfo ? <span style={{ color: "#5b8dee" }}>{skuInfo.nombre}</span>
                                   : prod.sku ? <span style={{ color: "#ff3b30" }}>No encontrado</span>
                                   : <span style={{ color: "#c7c7cc" }}>—</span>}
                                </td>
                                <td style={{ padding: "3px 6px" }}><input className="celda" type="number" step="0.01" value={prod.delta_precio} onChange={e => updateProd(filaC._id, paso._id, prod._id, "delta_precio", e.target.value)} placeholder="0" style={{ minWidth: 80, textAlign: "right" }} /></td>
                                <td style={{ textAlign: "center", padding: "4px 6px" }}><button onClick={() => removeProd(filaC._id, paso._id, prod._id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ff3b30" }}><X size={14} /></button></td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                      <button className="btn-addrow" style={{ margin: "6px 10px 10px" }} onClick={() => addProd(filaC._id, paso._id)}><Plus size={13} /> Agregar SKU</button>
                    </div>
                  ))}
                  <button className="btn-addrow" onClick={() => addPaso(filaC._id)}><Plus size={14} /> Agregar paso</button>
                </div>
                <div style={{ borderTop: "1px solid rgba(0,0,0,0.07)" }}>
                  {comboError && (
                    <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 22px", background: "rgba(255,59,48,0.07)", borderBottom: "1px solid rgba(255,59,48,0.12)" }}>
                      <AlertTriangle size={14} color="#ff3b30" />
                      <span style={{ fontSize: 13, color: "#ff3b30", fontWeight: 500 }}>{comboError}</span>
                    </div>
                  )}
                  <div style={{ padding: "14px 22px", display: "flex", alignItems: "center", gap: 8 }}>
                    <button className="btn-ghost" style={{ fontSize: 13 }} onClick={descargarTemplateCombo}><Download size={13} /> Plantilla</button>
                    <label style={{ display: "inline-flex" }}>
                      <span className="btn-ghost" style={{ fontSize: 13, cursor: "pointer" }}><Upload size={13} /> Importar Excel</span>
                      <input ref={comboImportRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => importarCombo(filaC._id, e)} />
                    </label>
                    <div style={{ flex: 1 }} />
                    <button className="btn-primary" style={{ fontSize: 14, padding: "9px 22px" }} onClick={() => {
                      if (hayOblig && !pctOk) { setComboError(`Los pasos obligatorios suman ${sumOblig}% — deben sumar exactamente 100%.`); return; }
                      setComboError(""); setComboModal(null);
                    }}>Listo</button>
                  </div>
                </div>
              </div>
            </div>
          );
        })()}
      </main>
    </AppShell>
  );
}

function VistaAdmin() {
  const [usuarios, setUsuarios] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [msj, setMsj] = useState(null);
  const [nuevoEmail, setNuevoEmail] = useState("");
  const [nuevoNombre, setNuevoNombre] = useState("");
  const [nuevoRol, setNuevoRol] = useState("solicitante");
  const [creando, setCreando] = useState(false);

  const cargar = async () => {
    setCargando(true);
    const { data } = await supabase.from("profiles").select("*").order("nombre");
    setUsuarios(data ?? []);
    setCargando(false);
  };

  useEffect(() => { cargar(); }, []);

  const actualizarRol = async (id, rol) => {
    const { error } = await supabase.from("profiles").update({ rol }).eq("id", id);
    if (!error) setUsuarios(u => u.map(x => x.id === id ? { ...x, rol } : x));
  };

  const toggleActivo = async (id, activo) => {
    const { error } = await supabase.from("profiles").update({ activo }).eq("id", id);
    if (!error) setUsuarios(u => u.map(x => x.id === id ? { ...x, activo } : x));
  };

  const crearUsuario = async (e) => {
    e.preventDefault();
    setCreando(true);
    setMsj(null);
    const { data, error } = await supabase.functions.invoke("invitar-usuario", {
      body: { email: nuevoEmail, nombre: nuevoNombre, rol: nuevoRol },
    });
    if (error || data?.error) {
      setMsj({ tipo: "error", texto: error?.message ?? data?.error });
    } else {
      setMsj({ tipo: "ok", texto: `Invitación enviada a ${nuevoEmail}. El usuario recibirá un email para crear su contraseña.` });
      setNuevoEmail(""); setNuevoNombre(""); setNuevoRol("solicitante");
      setTimeout(cargar, 1500);
    }
    setCreando(false);
  };

  return (
    <main className="sol-wrap">
      <div className="sol-head">
        <div>
          <h1 className="sol-title">Administración de Usuarios</h1>
          <p className="sol-sub">Gestiona accesos y roles del portal. Solo visible para el equipo Datos Maestros.</p>
        </div>
      </div>

      {/* Crear usuario */}
      <section className="seccion">
        <h2 className="ayuda-h2"><User size={20} /> Invitar nuevo usuario</h2>
        <form onSubmit={crearUsuario} className="admin-form">
          <div className="admin-form-row">
            <div className="login-field">
              <label>Nombre</label>
              <input value={nuevoNombre} onChange={e => setNuevoNombre(e.target.value)} placeholder="Nombre completo" required />
            </div>
            <div className="login-field">
              <label>Correo electrónico</label>
              <input type="email" value={nuevoEmail} onChange={e => setNuevoEmail(e.target.value)} placeholder="nombre@arcoprime.cl" required />
            </div>
            <div className="login-field">
              <label>Rol</label>
              <select value={nuevoRol} onChange={e => setNuevoRol(e.target.value)} className="admin-select">
                <option value="solicitante">Solicitante</option>
                <option value="datos_maestros">Datos Maestros</option>
              </select>
            </div>
            <button type="submit" className="btn-primary" disabled={creando} style={{ alignSelf: "flex-end" }}>
              {creando ? "Enviando…" : "Enviar invitación"}
            </button>
          </div>
          {msj && <p className={msj.tipo === "ok" ? "admin-msj-ok" : "login-error"} style={{ marginTop: 8 }}>{msj.tipo === "ok" ? <CheckCircle2 size={14} /> : <AlertTriangle size={14} />} {msj.texto}</p>}
        </form>
      </section>

      {/* Lista de usuarios */}
      <section className="seccion">
        <h2 className="ayuda-h2"><Layers size={20} /> Usuarios registrados</h2>
        {cargando ? (
          <p style={{ color: "#86868b", fontSize: 14 }}>Cargando usuarios…</p>
        ) : (
          <div className="admin-tabla-wrap">
            <table className="admin-tabla">
              <thead>
                <tr>
                  <th>Nombre</th>
                  <th>Correo</th>
                  <th>Rol</th>
                  <th>Activo</th>
                </tr>
              </thead>
              <tbody>
                {usuarios.map(u => (
                  <tr key={u.id}>
                    <td><strong>{u.nombre}</strong></td>
                    <td style={{ color: "#6e6e73" }}>{u.email}</td>
                    <td>
                      <select
                        value={u.rol}
                        onChange={e => actualizarRol(u.id, e.target.value)}
                        className={"admin-rol-sel " + (u.rol === "datos_maestros" ? "rol-ddmm" : "rol-sol")}
                      >
                        <option value="solicitante">Solicitante</option>
                        <option value="datos_maestros">Datos Maestros</option>
                      </select>
                    </td>
                    <td>
                      <button
                        onClick={() => toggleActivo(u.id, !u.activo)}
                        className={"admin-activo " + (u.activo ? "activo-on" : "activo-off")}
                      >
                        {u.activo ? "Activo" : "Inactivo"}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </main>
  );
}

/* ============================================================
   VISTA LOGIN
   ============================================================ */
function VistaLogin() {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [modo, setModo] = useState("login"); // 'login' | 'recovery'
  const [recoveryEnviado, setRecoveryEnviado] = useState(false);

  const login = async (e) => {
    e.preventDefault();
    setLoading(true); setError(null);
    const { error } = await supabase.auth.signInWithPassword({ email, password: pass });
    if (error) { setError("Correo o contraseña incorrectos."); setLoading(false); }
  };

  const enviarRecovery = async (e) => {
    e.preventDefault();
    setLoading(true); setError(null);
    const redirectTo = window.location.origin + window.location.pathname;
    const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo });
    setLoading(false);
    if (error) { setError("No se pudo enviar el correo. Verifica el email."); }
    else { setRecoveryEnviado(true); }
  };

  return (
    <div className="login-wrap">
      <div className="login-box">
        <div className="login-logo">NEXUS</div>

        {modo === "login" ? (
          <>
            <h1 className="login-title">Iniciar sesión</h1>
            <p className="login-sub">Portal de Solicitudes de Datos Maestros · Arcoprime</p>
            <form onSubmit={login} className="login-form">
              <div className="login-field">
                <label>Correo electrónico</label>
                <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="nombre@arcoprime.cl" required autoFocus />
              </div>
              <div className="login-field">
                <label>Contraseña</label>
                <input type="password" value={pass} onChange={e => setPass(e.target.value)} placeholder="••••••••" required />
              </div>
              {error && <p className="login-error"><AlertTriangle size={14} /> {error}</p>}
              <button type="submit" className="btn-primary" disabled={loading} style={{ marginTop: 4 }}>
                {loading ? "Ingresando…" : "Ingresar"}
              </button>
            </form>
            <button className="login-link" onClick={() => { setModo("recovery"); setError(null); }}>
              ¿Olvidaste tu contraseña?
            </button>
          </>
        ) : recoveryEnviado ? (
          <>
            <h1 className="login-title">Correo enviado</h1>
            <p className="login-sub">Revisa tu bandeja de entrada en <strong>{email}</strong>. El link expira en 1 hora.</p>
            <button className="login-link" onClick={() => { setModo("login"); setRecoveryEnviado(false); }}>
              ← Volver al inicio de sesión
            </button>
          </>
        ) : (
          <>
            <h1 className="login-title">Recuperar contraseña</h1>
            <p className="login-sub">Te enviaremos un link para crear una nueva contraseña.</p>
            <form onSubmit={enviarRecovery} className="login-form">
              <div className="login-field">
                <label>Correo electrónico</label>
                <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="nombre@arcoprime.cl" required autoFocus />
              </div>
              {error && <p className="login-error"><AlertTriangle size={14} /> {error}</p>}
              <button type="submit" className="btn-primary" disabled={loading} style={{ marginTop: 4 }}>
                {loading ? "Enviando…" : "Enviar link de recuperación"}
              </button>
            </form>
            <button className="login-link" onClick={() => { setModo("login"); setError(null); }}>
              ← Volver al inicio de sesión
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function VistaCrearContrasena({ onDone }) {
  const [pass, setPass] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [listo, setListo] = useState(false);

  const guardar = async (e) => {
    e.preventDefault();
    setError(null);
    if (pass.length < 6) { setError("La contraseña debe tener al menos 6 caracteres."); return; }
    if (pass !== confirm) { setError("Las contraseñas no coinciden."); return; }
    setLoading(true);
    const { error } = await supabase.auth.updateUser({ password: pass });
    if (error) { setError("No se pudo guardar. Intenta de nuevo."); setLoading(false); return; }
    setListo(true);
    setTimeout(onDone, 1800);
  };

  return (
    <div className="login-wrap">
      <div className="login-box">
        <div className="login-logo">NEXUS</div>
        <h1 className="login-title">Crear contraseña</h1>
        <p className="login-sub">Bienvenido/a al portal. Elige una contraseña para tu cuenta.</p>
        {listo ? (
          <p className="login-ok"><CheckCircle2 size={16} /> Contraseña guardada. Ingresando…</p>
        ) : (
          <form onSubmit={guardar} className="login-form">
            <div className="login-field">
              <label>Nueva contraseña</label>
              <input type="password" value={pass} onChange={e => setPass(e.target.value)} placeholder="Mínimo 6 caracteres" required autoFocus />
            </div>
            <div className="login-field">
              <label>Confirmar contraseña</label>
              <input type="password" value={confirm} onChange={e => setConfirm(e.target.value)} placeholder="Repite tu contraseña" required />
            </div>
            {error && <p className="login-error"><AlertTriangle size={14} /> {error}</p>}
            <button type="submit" className="btn-primary" disabled={loading} style={{ marginTop: 4 }}>
              {loading ? "Guardando…" : "Crear contraseña"}
            </button>
          </form>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   SHELL DE APLICACIÓN: sidebar + contenido + panel derecho
   ============================================================ */
function AppShell({ vista, setVista, perfil, children, rightPanel }) {
  return (
    <div className={"shell" + (rightPanel ? " shell-has-right" : "")}>
      <Estilos />
      <Sidebar vista={vista} setVista={setVista} perfil={perfil} />
      <div className="shell-body">
        <main className="shell-content">{children}</main>
        {rightPanel && <aside className="shell-right">{rightPanel}</aside>}
      </div>
    </div>
  );
}

function Sidebar({ vista, setVista, perfil }) {
  const [menuAbierto, setMenuAbierto] = useState(false);
  const cerrarSesion = async () => { await supabase.auth.signOut(); };
  const esDDMM = perfil?.rol === "datos_maestros";

  const SKU_IDS = ["nuevo-sku", "nueva-receta", "nueva-combo", "mod-receta", "mod-combo"];
  const skuActivo = SKU_IDS.includes(vista);
  const [skuAbierto, setSkuAbierto] = useState(skuActivo);

  const SKU_SUB = [
    { id: "nuevo-sku",    label: "Nuevo SKU",        icon: Package },
    { id: "nueva-receta", label: "Nueva Receta",      icon: Utensils },
    { id: "nueva-combo",  label: "Nuevo Combo",       icon: Gift },
    { id: "mod-receta",   label: "Modificar Receta",  icon: Pencil },
    { id: "mod-combo",    label: "Modificar Combo",   icon: Pencil },
  ];

  const TABS = [
    { id: "inicio",      label: "Inicio",            icon: Home },
    { id: "nueva",       label: "Nueva solicitud",   icon: FileText },
    { id: "solicitudes", label: "Solicitudes",       icon: Inbox },
    { id: "clusters",    label: "Gestor de locales", icon: Boxes },
    { id: "cvp",         label: "Ciclo de Vida",     icon: Layers },
    { id: "ayuda",       label: "Centro de Ayuda",   icon: Lightbulb },
    ...(esDDMM ? [
      { id: "maestros",  label: "Bases Maestras",    icon: Database },
      { id: "planillas", label: "Planillas",          icon: Table2 },
      { id: "admin",     label: "Admin",              icon: ShieldCheck },
    ] : []),
  ];

  const navegarA = id => { setVista(id); setMenuAbierto(false); };

  const grupoSKU = (
    <>
      <button
        className={"sidebar-item" + (skuActivo ? " on" : "")}
        onClick={() => setSkuAbierto(o => !o)}
        style={{ justifyContent: "space-between" }}
      >
        <span style={{ display: "flex", alignItems: "center", gap: 11 }}>
          <Package size={16} strokeWidth={1.8} /><span>Creación SKU</span>
        </span>
        {skuAbierto ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
      </button>
      {skuAbierto && (
        <div style={{ display: "flex", flexDirection: "column", gap: 1, marginLeft: 12, paddingLeft: 12, borderLeft: "2px solid rgba(255,255,255,0.1)" }}>
          {SKU_SUB.map(t => {
            const Icon = t.icon;
            return (
              <button key={t.id} className={"sidebar-item" + (vista === t.id ? " on" : "")}
                style={{ fontSize: 13, padding: "8px 12px" }}
                onClick={() => navegarA(t.id)}>
                <Icon size={14} strokeWidth={1.8} /><span>{t.label}</span>
              </button>
            );
          })}
        </div>
      )}
    </>
  );

  return (
    <>
      {/* Desktop sidebar */}
      <aside className="sidebar">
        <div className="sidebar-logo">
          <div>NEXUS</div>
          <div className="sidebar-subtitulo">Gestor de Solicitudes y Ciclo de Vida del Producto</div>
        </div>
        <nav className="sidebar-nav">
          <button key="inicio" className={"sidebar-item" + (vista === "inicio" ? " on" : "")} onClick={() => navegarA("inicio")}>
            <Home size={16} strokeWidth={1.8} /><span>Inicio</span>
          </button>
          {grupoSKU}
          {TABS.filter(t => t.id !== "inicio").map(t => {
            const Icon = t.icon;
            return (
              <button key={t.id} className={"sidebar-item" + (vista === t.id ? " on" : "")} onClick={() => navegarA(t.id)}>
                <Icon size={16} strokeWidth={1.8} /><span>{t.label}</span>
              </button>
            );
          })}
        </nav>
        <div className="sidebar-footer">
          {perfil && <div className="sidebar-user">{perfil.nombre}</div>}
          <button className="sidebar-logout" onClick={cerrarSesion}>Cerrar sesión</button>
        </div>
      </aside>

      {/* Mobile top bar */}
      <div className="topbar-mobile">
        <span className="nav-logo">NEXUS</span>
        <button className="nav-hamburger" onClick={() => setMenuAbierto(o => !o)} aria-label="Menú">
          {menuAbierto ? <X size={20} /> : <Menu size={20} />}
        </button>
      </div>

      {/* Mobile overlay */}
      {menuAbierto && (
        <div className="nav-mobile-overlay" onClick={() => setMenuAbierto(false)}>
          <div className="nav-mobile-menu" onClick={e => e.stopPropagation()}>
            {perfil && <div className="nav-mobile-user">{perfil.nombre}</div>}
            <button className={"nav-mobile-item" + (vista === "inicio" ? " on" : "")} onClick={() => navegarA("inicio")}>Inicio</button>
            <div style={{ fontSize: 11, fontWeight: 700, color: "rgba(255,255,255,0.4)", padding: "8px 16px 4px", textTransform: "uppercase", letterSpacing: ".06em" }}>Creación SKU</div>
            {SKU_SUB.map(t => (
              <button key={t.id} className={"nav-mobile-item" + (vista === t.id ? " on" : "")} onClick={() => navegarA(t.id)} style={{ paddingLeft: 28 }}>
                {t.label}
              </button>
            ))}
            {TABS.filter(t => t.id !== "inicio").map(t => (
              <button key={t.id} className={"nav-mobile-item" + (vista === t.id ? " on" : "")} onClick={() => navegarA(t.id)}>
                {t.label}
              </button>
            ))}
            <button className="nav-mobile-salir" onClick={cerrarSesion}>Cerrar sesión</button>
          </div>
        </div>
      )}
    </>
  );
}

function PanelValidacion({ plSel, totales, totalOk, totalErr, todoValidado, puedeEnviar, perfil, enviar, estado }) {
  const solicitante = perfil?.nombre || "";
  if (plSel.length === 0) {
    return (
      <div className="panel-vacio">
        <ShieldCheck size={38} strokeWidth={1.2} color="#b0b0b5" />
        <p>Selecciona un tipo de solicitud para comenzar</p>
      </div>
    );
  }

  const totalFilas = totales.reduce((a, t) => a + t.ok + t.err, 0);

  return (
    <div className="panel-val">
      <h2 className="panel-titulo"><ShieldCheck size={17} /> Validación</h2>

      {totalFilas > 0 && (
        <div className="panel-progreso">
          <span className="panel-prog-txt">{totalOk} de {totalFilas} filas válidas</span>
          <div className="panel-bar">
            <div className="panel-bar-fill" style={{ width: `${Math.round(totalOk / totalFilas * 100)}%`, background: totalErr > 0 ? "#ff9500" : "#34c759" }} />
          </div>
          {totalErr > 0 && <span className="panel-badge-err"><XCircle size={12} /> {totalErr} con error</span>}
        </div>
      )}

      {!todoValidado && (
        <div className="panel-lista">
          {plSel.map(p => {
            const t = totales.find(x => x.id === p.id);
            return (
              <div key={p.id} className={"panel-item" + (t?.validado ? (t.err > 0 ? " item-err" : " item-ok") : "")}>
                <p.icon size={14} /><span>{p.nombre}</span>
                <span className="panel-item-tag">
                  {!t?.validado ? "Pendiente" : t.err > 0 ? `${t.err} errores` : `${t.ok} OK`}
                </span>
              </div>
            );
          })}
        </div>
      )}

      {todoValidado && totalErr > 0 && (
        <div className="panel-errores">
          {plSel.flatMap(p => {
            const res = estado[p.id]?.resultado;
            if (!res) return [];
            return res.filas.filter(f => !f.ok).slice(0, 4).map((f, i) => (
              <div key={p.id + i} className="panel-err-item">
                <span className="panel-err-dot" />
                <div>
                  <div><strong>Fila {f.gi + 1}</strong> · {p.nombre}</div>
                  <div className="panel-err-txt">{f.errores[0]}</div>
                </div>
              </div>
            ));
          })}
        </div>
      )}

      <div className="panel-footer">
        <div className="panel-sol-row">
          <User size={14} color="#636366" />
          <span style={{ flex: 1, fontSize: 13, color: solicitante ? "#3a3a3c" : "#aeaeb2", fontWeight: solicitante ? 500 : 400 }}>
            {solicitante || "Sin usuario autenticado"}
          </span>
        </div>
        <p className="panel-hint">
          {!todoValidado ? "Valida cada planilla para continuar"
            : totalErr > 0 ? "Corrige los errores y vuelve a validar"
            : "Todo listo. Revisa y envía."}
        </p>
        <button className="btn-primary" style={{ width: "100%" }} disabled={!puedeEnviar} onClick={enviar}>
          Enviar solicitud <ArrowRight size={16} />
        </button>
      </div>
    </div>
  );
}

function NavBar({ info, vista, setVista, perfil }) {
  const [menuAbierto, setMenuAbierto] = useState(false);
  const cerrarSesion = async () => { await supabase.auth.signOut(); };
  const esDDMM = perfil?.rol === "datos_maestros";

  const TABS = [
    { id: "nueva",        label: "Nueva solicitud" },
    { id: "nueva-receta", label: "Nueva Receta" },
    { id: "nueva-combo",  label: "Nuevo Combo" },
    { id: "solicitudes",  label: "Solicitudes" },
    { id: "clusters",     label: "Gestor de locales" },
    { id: "cvp",          label: "Ciclo de Vida" },
    { id: "ayuda",        label: "Centro de Ayuda" },
    ...(esDDMM ? [{ id: "maestros", label: "Bases Maestras" }, { id: "admin", label: "Admin" }] : []),
  ];

  const navegarA = (id) => { setVista(id); setMenuAbierto(false); };

  return (
    <>
      <nav className="nav">
        <div className="nav-inner">
          <span className="nav-logo">NEXUS</span>

          {/* Tabs desktop */}
          <span className="nav-tabs nav-tabs-desktop">
            {TABS.map(t => (
              <button key={t.id} className={"nav-tab" + (vista === t.id ? " on" : "")} onClick={() => navegarA(t.id)}>{t.label}</button>
            ))}
          </span>

          <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span className={"nav-badge" + (info && !info.error ? " on" : "")}>
              <Database size={13} /> {info && !info.error ? "Base completa" : "Base demo"}
            </span>
            {perfil && <span className="nav-nombre">{perfil.nombre}</span>}
            <button className="nav-logout" onClick={cerrarSesion}>Salir</button>
            {/* Hamburger mobile */}
            <button className="nav-hamburger" onClick={() => setMenuAbierto(o => !o)} aria-label="Menú">
              {menuAbierto ? <X size={20} /> : <Menu size={20} />}
            </button>
          </span>
        </div>
      </nav>

      {/* Overlay mobile */}
      {menuAbierto && (
        <div className="nav-mobile-overlay" onClick={() => setMenuAbierto(false)}>
          <div className="nav-mobile-menu" onClick={e => e.stopPropagation()}>
            {perfil && <div className="nav-mobile-user">{perfil.nombre}</div>}
            {TABS.map(t => (
              <button key={t.id} className={"nav-mobile-item" + (vista === t.id ? " on" : "")} onClick={() => navegarA(t.id)}>
                {t.label}
              </button>
            ))}
            <button className="nav-mobile-salir" onClick={cerrarSesion}>Cerrar sesión</button>
          </div>
        </div>
      )}
    </>
  );
}

/* ============================================================
   VISTA SOLICITUDES: listado, detalle y gestión de estado
   ============================================================ */
function VistaSolicitudes({ perfil }) {
  const [cargando, setCargando] = useState(true);
  const [sols, setSols] = useState([]);
  const [abierta, setAbierta] = useState(null);
  const [filtro, setFiltro] = useState("Todas");
  const [rechazando, setRechazando] = useState(null); // folio en rechazo
  const [motivo, setMotivo] = useState("");
  const [notifMsg, setNotifMsg] = useState(null);
  const esDDMM = perfil?.rol === "datos_maestros";

  const recargar = useCallback(async () => {
    setCargando(true);
    const lista = await listarSolicitudes(perfil);
    setSols(lista);
    setCargando(false);
  }, [perfil]);

  React.useEffect(() => { recargar(); }, [recargar]);

  const notificar = async (sol, est, motivoRec) => {
    try {
      await supabase.functions.invoke("notificar-solicitud", {
        body: {
          folio: sol.folio,
          estado: est,
          solicitante_nombre: sol.solicitante_nombre,
          solicitante_email: sol.solicitante_email,
          planillas: sol.planillas,
          motivo_rechazo: motivoRec || null,
        },
      });
      if (est === "Aprobada") {
        const attachments = (sol.planillas || []).map(p => {
          const plConfig = PLANILLAS.find(x => x.id === p.id);
          if (!plConfig) return null;
          const aoa = [plConfig.salidaLabels];
          p.filas.forEach(f => aoa.push(plConfig.salida.map(k => f.data[k] ?? f.auto[k] ?? "")));
          const ws = XLSX.utils.aoa_to_sheet(aoa);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Datos");
          const base64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
          return { base64, filename: `${p.id}_SAP.xlsx` };
        }).filter(Boolean);

        await supabase.functions.invoke("crear-ticket-freshdesk", {
          body: {
            folio: sol.folio,
            solicitante_nombre: sol.solicitante_nombre,
            solicitante_email: sol.solicitante_email,
            planillas: sol.planillas,
            attachments,
          },
        });
      }
    } catch (e) { console.error("notificar:", e); }
  };

  const setEstadoSol = async (folio, est, motivoRec) => {
    const actualizada = await cambiarEstado(folio, est, motivoRec);
    if (!actualizada) return;
    setSols(s => s.map(x => x.folio === folio ? actualizada : x));
    if (est === "En proceso" || est === "Aprobada" || est === "Rechazada") {
      notificar(actualizada, est, motivoRec);
      setNotifMsg(
        est === "Aprobada"   ? "✓ Solicitud aprobada — se notificó al solicitante y se creó el ticket en Freshdesk." :
        est === "Rechazada"  ? "✓ Solicitud rechazada — se notificó al solicitante." :
                               "✓ Solicitud en proceso — se notificó al solicitante."
      );
      setTimeout(() => setNotifMsg(null), 5000);
    }
  };

  const confirmarRechazo = async (folio) => {
    if (!motivo.trim()) return;
    await setEstadoSol(folio, "Rechazada", motivo.trim());
    setRechazando(null);
    setMotivo("");
  };

  const exportarDeSolicitud = (plData) => {
    const pl = PLANILLAS.find(p => p.id === plData.id);
    if (pl) exportarSalida(pl, plData.filas);
  };

  const visibles = filtro === "Todas" ? sols : sols.filter(s => s.estado === filtro);

  return (
    <main className="sol-wrap">
      <div className="sol-head">
        <div>
          <h1 className="sol-title">Solicitudes</h1>
          <p className="sol-sub">{esDDMM ? "Todas las solicitudes del portal. Aprueba o rechaza para notificar al solicitante." : "Tus solicitudes enviadas. Te notificaremos por email cuando cambien de estado."}</p>
        </div>
        <button className="btn-ghost" onClick={recargar}><RotateCcw size={14} /> Actualizar</button>
      </div>

      {notifMsg && <div className="sol-notif">{notifMsg}</div>}

      <div className="sol-filtros">
        {["Todas", ...Object.keys(ESTADOS)].map(f => (
          <button key={f} className={"chip" + (filtro === f ? " on" : "")} onClick={() => setFiltro(f)}>
            {f}{f !== "Todas" && <span className="chip-n">{sols.filter(s => s.estado === f).length}</span>}
          </button>
        ))}
      </div>

      {cargando ? (
        <div className="sol-vacio"><Clock size={28} strokeWidth={1.4} /><p>Cargando solicitudes…</p></div>
      ) : visibles.length === 0 ? (
        <div className="sol-vacio">
          <Inbox size={32} strokeWidth={1.3} />
          <p>{sols.length === 0 ? "Aún no hay solicitudes enviadas." : "No hay solicitudes con este estado."}</p>
        </div>
      ) : (
        <div className="sol-lista">
          {visibles.map(s => {
            const open = abierta === s.folio;
            const est = ESTADOS[s.estado] || ESTADOS["Enviada"];
            return (
              <div key={s.folio} className="sol-card">
                <button className="sol-fila" onClick={() => setAbierta(open ? null : s.folio)}>
                  <div className="sol-folio">
                    <strong>{s.folio}</strong>
                    <span>{fmtFecha(s.created_at || s.fecha)} · {s.solicitante_nombre || "Sin nombre"}</span>
                  </div>
                  <div className="sol-tipos">
                    {(s.planillas || []).map(p => {
                      const Pl = PLANILLAS.find(x => x.id === p.id);
                      const Ic = Pl?.icon;
                      return <span key={p.id} className="sol-tipo">{Ic && <Ic size={13} />}{p.nombre} · {p.filas.length}</span>;
                    })}
                  </div>
                  <span className="sol-estado" style={{ color: est.color, background: est.bg }}>{s.estado}</span>
                  {open ? <ChevronDown size={17} /> : <ChevronRight size={17} />}
                </button>

                {open && (
                  <div className="sol-detalle">
                    {esDDMM && (
                      <div className="sol-estado-edit">
                        <span>Cambiar estado:</span>
                        {["En proceso", "Aprobada"].map(e => (
                          <button key={e} className={"chip mini" + (s.estado === e ? " on" : "")}
                            onClick={() => setEstadoSol(s.folio, e)}>{e}</button>
                        ))}
                        {rechazando === s.folio ? (
                          <div className="sol-rechazo-form">
                            <input
                              className="sol-rechazo-input"
                              placeholder="Motivo del rechazo..."
                              value={motivo}
                              onChange={e => setMotivo(e.target.value)}
                              onKeyDown={e => e.key === "Enter" && confirmarRechazo(s.folio)}
                              autoFocus
                            />
                            <button className="btn-primary" style={{ fontSize: 13, padding: "6px 14px" }} onClick={() => confirmarRechazo(s.folio)}>Confirmar</button>
                            <button className="btn-ghost" onClick={() => { setRechazando(null); setMotivo(""); }}>Cancelar</button>
                          </div>
                        ) : (
                          <button className={"chip mini" + (s.estado === "Rechazada" ? " on" : "")}
                            style={{ color: "#c2271c" }}
                            onClick={() => { setRechazando(s.folio); setMotivo(""); }}>Rechazar</button>
                        )}
                        <button className={"chip mini" + (s.estado === "Aplicada" ? " on" : "")}
                          onClick={() => setEstadoSol(s.folio, "Aplicada")}>Aplicada</button>
                      </div>
                    )}
                    {s.motivo_rechazo && (
                      <div className="sol-motivo-rechazo">
                        <strong>Motivo rechazo:</strong> {s.motivo_rechazo}
                      </div>
                    )}

                    {(s.planillas || []).map(p => {
                      const pl = PLANILLAS.find(x => x.id === p.id);
                      if (!pl) return null;
                      return (
                        <div key={p.id} className="tabla-wrap">
                          <div className="tabla-head">
                            <span><strong>{p.nombre}</strong> · {p.filas.length} filas</span>
                            <button className="link-export" onClick={() => exportarDeSolicitud(p)}><Download size={13} /> Excel</button>
                          </div>
                          <div className="tabla-scroll">
                            <table>
                              <thead>
                                <tr>
                                  {pl.userCols.map(c => <th key={c.key}>{c.label}</th>)}
                                  {pl.autoCols.map(c => <th key={c.key} className="th-auto">{c.label}</th>)}
                                </tr>
                              </thead>
                              <tbody>
                                {p.filas.map((f, i) => (
                                  <tr key={i}>
                                    {pl.userCols.map(c => <td key={c.key}>{norm(f.data[c.key]) || "—"}</td>)}
                                    {pl.autoCols.map(c => <td key={c.key} className="td-auto">{f.auto[c.key]}</td>)}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </main>
  );
}

function Estilos() {
  return (
    <style>{`
      * { box-sizing: border-box; margin: 0; }
      .portal { min-height: 100vh; background: #e4e7f1; color: #1d1d1f;
        font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        -webkit-font-smoothing: antialiased; }

      /* ── Shell layout ── */
      .shell { display: flex; min-height: 100vh; background: #f0f2f8;
        font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        -webkit-font-smoothing: antialiased; color: #1d1d1f; }
      .shell-body { flex: 1; display: flex; min-width: 0; }
      .shell-content { flex: 1; min-width: 0; overflow: hidden; padding: 32px 28px; }
      .shell-right { width: 300px; flex-shrink: 0; border-left: 1px solid rgba(0,0,0,0.07);
        background: #fff; position: sticky; top: 0; height: 100vh; overflow-y: auto;
        display: flex; flex-direction: column; }

      /* ── Sidebar ── */
      .sidebar { width: 264px; flex-shrink: 0; background: #1a1a2e; color: #fff;
        display: flex; flex-direction: column; padding: 0; position: sticky; top: 0;
        height: 100vh; overflow-y: auto; }
      .sidebar-logo { font-weight: 700; font-size: 17px; letter-spacing: -0.02em;
        padding: 28px 22px 22px; border-bottom: 1px solid rgba(255,255,255,0.08); }
      .sidebar-nav { flex: 1; padding: 14px 12px; display: flex; flex-direction: column; gap: 3px; }
      .sidebar-item { display: flex; align-items: center; gap: 11px; padding: 11px 14px;
        border-radius: 11px; border: none; background: none; color: rgba(255,255,255,0.6);
        font: inherit; font-size: 15px; font-weight: 500; cursor: pointer; text-align: left;
        width: 100%; transition: all .15s; }
      .sidebar-item:hover { background: rgba(255,255,255,0.08); color: rgba(255,255,255,0.9); }
      .sidebar-item.on { background: rgba(255,255,255,0.13); color: #fff; font-weight: 600; }
      .sidebar-footer { border-top: 1px solid rgba(255,255,255,0.08); padding: 18px 16px; }
      .sidebar-subtitulo { font-size: 11px; font-weight: 400; color: rgba(255,255,255,0.4); margin-top: 4px; line-height: 1.4; }
      .content-header { margin-bottom: 24px; }
      .content-title { font-size: 26px; font-weight: 700; letter-spacing: 0.04em; color: #1d1d1f; text-transform: uppercase; }
      .content-sub { font-size: 14px; color: #86868b; margin-top: 3px; text-transform: capitalize; }
      .sidebar-user { font-size: 13px; color: rgba(255,255,255,0.45); padding: 0 6px 10px;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .sidebar-logout { width: 100%; background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.12);
        color: rgba(255,255,255,0.7); border-radius: 8px; padding: 8px; font: inherit; font-size: 14px;
        cursor: pointer; transition: all .15s; }
      .sidebar-logout:hover { background: rgba(255,59,48,0.15); border-color: rgba(255,59,48,0.3); color: #ff6b6b; }
      .topbar-mobile { display: none; }

      /* ── Panel derecho (validación) ── */
      .panel-vacio { flex: 1; display: flex; flex-direction: column; align-items: center;
        justify-content: center; gap: 14px; padding: 32px; text-align: center;
        color: #86868b; font-size: 15px; }
      .panel-val { display: flex; flex-direction: column; gap: 0; height: 100%; }
      .panel-titulo { font-size: 15px; font-weight: 700; color: #1d1d1f; padding: 22px 20px 16px;
        border-bottom: 1px solid rgba(0,0,0,0.06); display: flex; align-items: center; gap: 8px; }
      .panel-progreso { padding: 16px 20px; border-bottom: 1px solid rgba(0,0,0,0.05); }
      .panel-prog-txt { font-size: 14px; font-weight: 600; color: #1d1d1f; display: block; margin-bottom: 8px; }
      .panel-bar { height: 6px; background: #e8e8ed; border-radius: 3px; overflow: hidden; margin-bottom: 8px; }
      .panel-bar-fill { height: 100%; border-radius: 3px; transition: width .4s; }
      .panel-badge-err { display: inline-flex; align-items: center; gap: 5px; font-size: 13px;
        font-weight: 600; color: #c2271c; background: rgba(255,59,48,0.08); padding: 3px 9px; border-radius: 980px; }
      .panel-lista { padding: 12px 16px; border-bottom: 1px solid rgba(0,0,0,0.05); display: flex; flex-direction: column; gap: 6px; }
      .panel-item { display: flex; align-items: center; gap: 8px; font-size: 14px; color: #515154;
        padding: 8px 10px; background: #f5f5f7; border-radius: 8px; }
      .panel-item span:nth-child(2) { flex: 1; font-weight: 500; }
      .panel-item-tag { font-size: 11px; font-weight: 600; color: #86868b; background: rgba(0,0,0,0.06);
        padding: 2px 8px; border-radius: 980px; white-space: nowrap; }
      .panel-item.item-ok .panel-item-tag { color: #248a3d; background: rgba(52,199,89,0.12); }
      .panel-item.item-err .panel-item-tag { color: #c2271c; background: rgba(255,59,48,0.1); }
      .panel-errores { padding: 12px 16px; flex: 1; overflow-y: auto; display: flex; flex-direction: column; gap: 8px; }
      .panel-err-item { display: flex; gap: 10px; font-size: 13.5px; color: #515154;
        background: rgba(255,59,48,0.04); border: 1px solid rgba(255,59,48,0.12);
        border-radius: 8px; padding: 10px 12px; }
      .panel-err-dot { width: 7px; height: 7px; border-radius: 50%; background: #ff3b30;
        flex-shrink: 0; margin-top: 4px; }
      .panel-err-txt { color: #86868b; font-size: 11.5px; margin-top: 2px; }
      .panel-footer { padding: 16px 20px; border-top: 1px solid rgba(0,0,0,0.06);
        display: flex; flex-direction: column; gap: 10px; margin-top: auto; }
      .panel-sol-row { display: flex; align-items: center; gap: 8px; background: #f5f5f7;
        border-radius: 10px; padding: 10px 12px; color: #86868b; }
      .panel-hint { font-size: 13px; color: #86868b; line-height: 1.4; }
      .inp-solicitante { background: none; border: none; outline: none; font: inherit;
        font-size: 14.5px; color: #1d1d1f; }
      .inp-solicitante::placeholder { color: #b0b0b5; }

      /* ── NavBar legacy (mantener para vistas que aún lo usen) ── */
      .nav { position: sticky; top: 0; z-index: 50; background: rgba(228,231,241,0.92);
        backdrop-filter: saturate(200%) blur(24px); -webkit-backdrop-filter: saturate(200%) blur(24px);
        border-bottom: 1px solid rgba(180,185,220,0.35); }
      .nav-inner { max-width: 1060px; margin: 0 auto; padding: 0 16px; height: 52px;
        display: flex; align-items: center; justify-content: space-between; font-size: 14px; gap: 10px; }
      .nav-logo { font-weight: 700; letter-spacing: -0.02em; font-size: 15px; white-space: nowrap; flex-shrink: 0; }
      .nav-tabs { display: flex; gap: 2px; background: rgba(180,185,220,0.25); border-radius: 980px; padding: 3px;
        overflow-x: auto; scrollbar-width: none; -ms-overflow-style: none; flex-shrink: 1; min-width: 0; }
      .nav-tabs::-webkit-scrollbar { display: none; }
      .nav-tab { border: none; background: none; cursor: pointer; font: inherit; font-size: 13.5px;
        font-weight: 500; color: #5a5f7a; padding: 5px 12px; border-radius: 980px;
        transition: background .2s, color .2s; white-space: nowrap; flex-shrink: 0; }
      .nav-tab.on { background: #fff; color: #1d1d1f; box-shadow: 0 1px 4px rgba(100,110,180,0.18); font-weight: 600; }

      /* VISTA SOLICITUDES */
      .sol-wrap { max-width: none; padding: 40px 32px 60px; }
      .sol-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }
      .sol-title { font-size: 36px; font-weight: 800; letter-spacing: 0.04em; color: #14162a; text-transform: uppercase; }
      .sol-sub { color: #6e6e80; font-size: 16px; margin-top: 6px; max-width: 560px; line-height: 1.55; text-transform: capitalize; }
      .sol-filtros { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 20px; }
      .chip { display: inline-flex; align-items: center; gap: 6px; border: 1px solid rgba(180,185,220,0.4);
        background: rgba(255,255,255,0.7); cursor: pointer; font: inherit; font-size: 14px; font-weight: 500; color: #5a5f7a;
        padding: 7px 14px; border-radius: 980px; transition: all .2s; backdrop-filter: blur(8px); }
      .chip.on { background: #5b8dee; color: #fff; border-color: #5b8dee; box-shadow: 0 4px 12px rgba(91,141,238,0.3); }
      .chip.mini { padding: 4px 11px; font-size: 13px; }
      .chip-n { font-size: 11px; opacity: .65; }
      .sol-vacio { display: flex; flex-direction: column; align-items: center; gap: 12px;
        color: #86868b; padding: 70px 20px; text-align: center; font-size: 15.5px; }
      .sol-lista { display: flex; flex-direction: column; gap: 12px; }
      .sol-card { background: #fff; border: 1px solid rgba(200,205,230,0.5); border-radius: 18px; overflow: hidden;
        box-shadow: 0 2px 12px rgba(100,110,180,0.07), 0 8px 32px rgba(100,110,180,0.05); }
      .sol-fila { width: 100%; display: flex; align-items: center; gap: 16px; padding: 16px 20px;
        background: none; border: none; cursor: pointer; font: inherit; color: inherit; text-align: left; flex-wrap: wrap; }
      .sol-fila:hover { background: rgba(228,231,241,0.4); }
      .sol-folio { display: flex; flex-direction: column; gap: 2px; min-width: 190px; }
      .sol-folio strong { font-size: 16px; font-weight: 600; letter-spacing: -0.01em; }
      .sol-folio span { font-size: 13.5px; color: #86868b; }
      .sol-tipos { flex: 1; display: flex; gap: 8px; flex-wrap: wrap; }
      .sol-tipo { display: inline-flex; align-items: center; gap: 5px; font-size: 13px; font-weight: 500;
        color: #515154; background: #f5f5f7; border-radius: 980px; padding: 4px 11px; }
      .sol-estado { font-size: 13px; font-weight: 600; border-radius: 980px; padding: 5px 12px; flex: none; }
      .sol-detalle { border-top: 1px solid rgba(0,0,0,0.06); padding: 16px 20px 20px; display: flex; flex-direction: column; gap: 14px; }
      .sol-estado-edit { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; font-size: 14px; color: #6e6e73; }
      .sol-notif { background: rgba(52,199,89,0.12); color: #248a3d; border-radius: 12px; padding: 12px 18px; font-size: 14px; font-weight: 500; margin-bottom: 16px; }
      .sol-rechazo-form { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-top: 4px; }
      .sol-rechazo-input { flex: 1; min-width: 220px; padding: 7px 12px; border: 1.5px solid #ff3b30; border-radius: 8px; font: inherit; font-size: 14px; outline: none; }
      .sol-motivo-rechazo { font-size: 14px; color: #c2271c; background: rgba(255,59,48,0.07); border-radius: 10px; padding: 10px 14px; margin-bottom: 8px; }
      .success-botones { display: flex; gap: 12px; justify-content: center; flex-wrap: wrap; }

      /* GUÍA DE SOLICITUDES */
      .guia-card { background: #fff; border: 1px solid rgba(200,205,230,0.5); border-radius: 18px; overflow: hidden;
        box-shadow: 0 2px 12px rgba(100,110,180,0.06); }
      .guia-card.open { box-shadow: 0 8px 32px rgba(100,110,180,0.12); }
      .guia-ic { background: linear-gradient(135deg, rgba(94,92,230,0.14), rgba(191,90,242,0.14)) !important; color: #5e5ce6 !important; }
      .guia-body { padding: 4px 20px 22px 20px; }
      .guia-body .clu-input { margin-bottom: 14px; }
      .guia-lista { display: flex; flex-direction: column; gap: 6px; max-height: 340px; overflow-y: auto; }
      .guia-item { display: flex; align-items: center; gap: 14px; text-align: left; background: rgba(228,231,241,0.3);
        border: 1px solid rgba(0,0,0,0.06); border-radius: 12px; padding: 12px 14px; cursor: pointer;
        font: inherit; color: inherit; transition: border-color .2s, background .2s; }
      .guia-item:hover { border-color: rgba(94,92,230,0.5); background: #fff; }
      .guia-item.on { border-color: #5e5ce6; box-shadow: 0 0 0 1px #5e5ce6; background: #fff; }
      .guia-num { flex: none; width: 26px; height: 26px; border-radius: 50%; background: #ececf3; color: #5e5ce6;
        font-size: 13px; font-weight: 700; display: flex; align-items: center; justify-content: center; }
      .guia-item-txt { flex: 1; display: flex; flex-direction: column; gap: 1px; min-width: 0; }
      .guia-item-txt strong { font-size: 15px; font-weight: 600; }
      .guia-item-txt em { font-style: normal; font-size: 13px; color: #86868b; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .guia-letras { flex: none; font-size: 11px; font-weight: 600; color: #b25dd6; font-family: ui-monospace, "SF Mono", Menlo, monospace; }
      .guia-detalle { margin-top: 16px; border: 1px solid rgba(94,92,230,0.25); background: rgba(94,92,230,0.04);
        border-radius: 14px; padding: 18px 20px; }
      .guia-detalle h3 { font-size: 17px; font-weight: 600; letter-spacing: -0.01em; }
      .guia-obj { font-size: 14.5px; color: #515154; margin-top: 4px; line-height: 1.5; }
      .guia-nota { color: #b25000; }
      .guia-planillas { display: flex; flex-direction: column; gap: 6px; margin: 14px 0; }
      .guia-row { display: flex; align-items: center; gap: 12px; min-height: 46px; background: #fff;
        border: 1px solid rgba(0,0,0,0.07); border-radius: 12px; padding: 8px 12px; }
      .guia-row.portal { border-color: rgba(52,199,89,0.35); background: rgba(52,199,89,0.05); }
      .guia-badge { flex: none; width: 46px; height: 28px; border-radius: 8px;
        background: #ececf3; color: #5e5ce6; font-size: 13px; font-weight: 700;
        font-family: ui-monospace, "SF Mono", Menlo, monospace;
        display: inline-flex; align-items: center; justify-content: center; }
      .guia-row.portal .guia-badge { background: rgba(52,199,89,0.18); color: #1d6e30; }
      .guia-row-nombre { flex: 1; font-size: 14.5px; color: #1d1d1f; line-height: 1.3; min-width: 0; }
      .guia-tag { flex: none; display: inline-flex; align-items: center; gap: 4px; font-size: 11px;
        font-weight: 600; color: #86868b; background: #f5f5f7; border-radius: 980px; padding: 4px 10px; white-space: nowrap; }
      .guia-tag.on { color: #248a3d; background: rgba(52,199,89,0.14); }
      .guia-tag.on svg { color: #34c759; }
      .guia-accion { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; margin-top: 4px; }
      .guia-resto { font-size: 13.5px; color: #86868b; }
      .guia-resto.solo { display: block; margin-top: 8px; }

      /* CENTRO DE AYUDA */
      .ayuda-h2 { display: flex; align-items: center; gap: 10px; font-size: 18px; font-weight: 600; letter-spacing: -0.01em; margin-bottom: 14px; color: #1d1d1f; }
      .ayuda-h2 svg { color: #5b8dee; }
      .ayuda-intro { font-size: 15px; color: #515154; margin-bottom: 18px; line-height: 1.5; }
      
      /* FLUJO DE PROCESAMIENTO */
      .flujo-pasos { display: flex; flex-direction: column; gap: 0; }
      .paso { display: flex; align-items: flex-start; gap: 16px; background: #fff; border: 1px solid rgba(0,0,0,0.08);
        border-radius: 14px; padding: 16px 18px; }
      .paso-1, .paso-2, .paso-3, .paso-4 { border-bottom-left-radius: 0; border-bottom-right-radius: 0; border-bottom: none; }
      .paso-5 { border-top-left-radius: 0; border-top-right-radius: 0; }
      .paso-num { flex: none; width: 32px; height: 32px; border-radius: 50%; background: linear-gradient(135deg, #5b8dee, #5e5ce6);
        color: #fff; font-size: 15px; font-weight: 700; display: flex; align-items: center; justify-content: center; }
      .paso-txt { flex: 1; display: flex; flex-direction: column; gap: 2px; }
      .paso-txt strong { font-size: 15px; font-weight: 600; color: #1d1d1f; }
      .paso-txt span { font-size: 14px; color: #6e6e73; line-height: 1.4; }
      .paso-arrow { align-self: center; color: #c7c7cc; margin: -4px 0; }
      
      /* CATÁLOGO DE PLANILLAS */
      .ayuda-lista { display: flex; flex-direction: column; gap: 6px; }
      .ayuda-item { background: #fff; border: 1px solid rgba(200,205,230,0.5); border-radius: 14px; overflow: hidden;
        box-shadow: 0 2px 8px rgba(100,110,180,0.05); }
      .ayuda-item.open { box-shadow: 0 6px 24px rgba(100,110,180,0.1); }
      .ayuda-item-btn { width: 100%; display: flex; align-items: center; gap: 12px; padding: 14px 16px; background: none;
        border: none; font: inherit; color: inherit; cursor: pointer; transition: background .15s; text-align: left; }
      .ayuda-item-btn:hover { background: rgba(228,231,241,0.4); }
      .ayuda-item.open .ayuda-item-btn { background: rgba(228,231,241,0.3); }
      .ayuda-item-ic { flex: none; display: flex; color: #5b8dee; }
      .ayuda-item-head { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 2px; }
      .ayuda-item-head strong { font-size: 16px; font-weight: 600; color: #1d1d1f; }
      .ayuda-item-head em { font-style: normal; font-size: 14px; color: #6e6e73; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .ayuda-item-toggle { flex: none; color: #c7c7cc; }
      
      .ayuda-item-body { padding: 16px; border-top: 1px solid rgba(200,205,230,0.4); background: rgba(228,231,241,0.2); }
      .ayuda-detalle { display: flex; flex-direction: column; gap: 8px; margin-bottom: 14px; }
      .ayuda-detalle:last-child { margin-bottom: 0; }
      .ayuda-detalle h4 { font-size: 14px; font-weight: 700; color: #1d1d1f; text-transform: uppercase; letter-spacing: 0.5px; display: flex; align-items: center; gap: 6px; }
      .ayuda-detalle p { font-size: 14px; color: #515154; line-height: 1.5; }
      
      .ayuda-cols { display: flex; flex-direction: column; gap: 6px; }
      .ayuda-col { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; padding: 8px 10px; background: #fff;
        border-radius: 10px; border: 1px solid rgba(0,0,0,0.06); font-size: 13px; }
      .col-nombre { flex: 1; font-weight: 600; color: #1d1d1f; min-width: 120px; }
      .col-tipo { background: rgba(91,141,238,0.1); color: #5b8dee; border-radius: 6px; padding: 2px 7px; font-weight: 600; }
      .col-hint { color: #86868b; font-family: ui-monospace, "SF Mono", Menlo, monospace; }
      .col-def { background: rgba(52,199,89,0.1); color: #248a3d; border-radius: 6px; padding: 2px 7px; font-size: 11px; }
      .col-opt { background: rgba(255,159,64,0.1); color: #c65d07; border-radius: 6px; padding: 2px 7px; font-size: 11px; }
      
      .ayuda-auto { display: flex; flex-wrap: wrap; gap: 6px; }
      .auto-item { display: inline-flex; align-items: center; gap: 4px; font-size: 13px; color: #248a3d;
        background: rgba(52,199,89,0.1); border-radius: 980px; padding: 4px 10px; }
      .auto-item.sin { color: #86868b; background: rgba(0,0,0,0.06); }
      
      .ayuda-ejemplo { display: flex; flex-direction: column; gap: 6px; padding: 10px; background: #fff; border-radius: 10px;
        border: 1px solid rgba(0,0,0,0.06); }
      .ej-fila { display: flex; gap: 10px; font-size: 13px; }
      .ej-etiq { font-weight: 600; color: #1d1d1f; min-width: 140px; }
      .ej-val { color: #515154; font-family: ui-monospace, "SF Mono", Menlo, monospace; }
      
      .ayuda-salida { display: flex; flex-wrap: wrap; gap: 6px; }
      .salida-col { display: inline-block; font-size: 13px; color: #515154; background: #fff; border: 1px solid rgba(0,0,0,0.08);
        border-radius: 8px; padding: 5px 10px; }
      
      /* VALIDACIONES */
      .ayuda-validaciones { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 14px; }
      .val-item { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 14px; padding: 16px; }
      .val-item h4 { display: flex; align-items: center; gap: 8px; font-size: 15px; font-weight: 600; margin-bottom: 8px; color: #1d1d1f; }
      .val-item h4 svg { color: #34c759; }
      .val-item p { font-size: 14px; color: #515154; line-height: 1.5; }
      
      /* CONSEJOS */
      .ayuda-consejos { display: flex; flex-direction: column; gap: 10px; list-style: none; padding: 0; margin: 0; }
      .ayuda-consejos li { font-size: 14.5px; color: #515154; line-height: 1.6; padding-left: 24px; position: relative; background: rgba(228,231,241,0.3);
        border-left: 3px solid #5b8dee; border-radius: 0 8px 8px 0; padding: 11px 14px 11px 18px; }
      .ayuda-consejos strong { color: #1d1d1f; font-weight: 600; }

      /* GESTOR DE LOCALES */
      .clu-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)); gap: 18px; }
      .clu-card { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 18px; padding: 22px; }
      .clu-card-head { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
      .clu-card-head h2 { font-size: 18px; font-weight: 600; letter-spacing: -0.01em; }
      .clu-card-head svg { color: #5b8dee; }
      .clu-help { font-size: 14.5px; color: #6e6e73; line-height: 1.5; margin-bottom: 14px; }
      .clu-textarea { width: 100%; border: 1px solid rgba(0,0,0,0.12); border-radius: 12px; padding: 12px 14px;
        font: inherit; font-size: 14.5px; resize: vertical; background: #f8f8fc; transition: border-color .15s, box-shadow .15s; }
      .clu-textarea:focus { outline: none; border-color: #5b8dee; box-shadow: 0 0 0 3px rgba(91,141,238,0.15); background: #fff; }
      .clu-textarea.sm { min-height: 60px; }
      .clu-input { width: 100%; border: 1px solid rgba(0,0,0,0.12); border-radius: 12px; padding: 11px 14px;
        font: inherit; font-size: 14.5px; background: #f8f8fc; transition: border-color .15s, box-shadow .15s; }
      .clu-input:focus { outline: none; border-color: #5b8dee; box-shadow: 0 0 0 3px rgba(91,141,238,0.15); background: #fff; }
      .clu-resultado { display: flex; align-items: center; gap: 10px; margin-top: 12px;
        background: #f5f5f7; border-radius: 12px; padding: 10px 14px; }
      .clu-linea { flex: 1; font-size: 13.5px; font-family: ui-monospace, "SF Mono", Menlo, monospace;
        color: #1d1d1f; word-break: break-all; max-height: 76px; overflow-y: auto; }
      .clu-conteo { display: flex; gap: 14px; flex-wrap: wrap; margin-top: 10px; font-size: 13.5px; color: #515154; }
      .clu-conteo span { display: inline-flex; align-items: center; gap: 5px; }
      .clu-card .btn-ghost { margin-top: 12px; }
      .clu-form { display: flex; flex-direction: column; gap: 10px; margin-bottom: 18px; }
      .clu-form-btns { display: flex; gap: 8px; flex-wrap: wrap; }
      .clu-vacio { font-size: 14.5px; color: #86868b; text-align: center; padding: 22px 10px;
        border: 1px dashed rgba(0,0,0,0.12); border-radius: 12px; }
      .clu-lista { display: flex; flex-direction: column; gap: 8px; }
      .clu-item { display: flex; align-items: center; gap: 12px; border: 1px solid rgba(0,0,0,0.07);
        border-radius: 12px; padding: 11px 14px; }
      .clu-item-info { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 2px; }
      .clu-item-info strong { font-size: 15px; font-weight: 600; }
      .clu-item-info span { font-size: 13px; color: #6e6e73; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .clu-item-info em { font-style: normal; font-weight: 600; }
      .clu-tip { display: flex; align-items: flex-start; gap: 10px; margin-top: 20px; font-size: 14px;
        color: #515154; background: rgba(91,141,238,0.07); border-radius: 14px; padding: 14px 18px; line-height: 1.5; }
      .clu-tip svg { color: #5b8dee; flex: none; margin-top: 1px; }
      .clu-tip code { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 6px;
        padding: 1px 6px; font-size: 13px; font-family: ui-monospace, "SF Mono", Menlo, monospace; }
      .nav-badge { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 500;
        color: #86868b; background: #f0f0f2; border-radius: 980px; padding: 5px 11px; }
      .nav-badge.on { color: #248a3d; background: rgba(52,199,89,0.12); }

      .hero { max-width: 780px; margin: 0 auto; padding: 76px 22px 48px; text-align: center; }
      .hero-eyebrow { display: inline-flex; align-items: center; gap: 6px; font-size: 14px; font-weight: 600;
        color: #5b8dee; background: rgba(91,141,238,0.1); border-radius: 980px; padding: 6px 14px; margin-bottom: 22px; }
      .hero-title { font-size: clamp(34px, 6vw, 54px); line-height: 1.07; font-weight: 700; letter-spacing: -0.018em; }
      .hero-grad { background: linear-gradient(90deg, #5b8dee, #5e5ce6 60%, #bf5af2);
        -webkit-background-clip: text; background-clip: text; color: transparent; }
      .hero-sub { margin: 22px auto 0; font-size: 17px; line-height: 1.55; color: #515154; max-width: 640px; }

      .seccion { max-width: 1060px; margin: 0 auto; padding: 20px 22px 34px; }
      .paso-head { display: flex; gap: 16px; align-items: flex-start; margin-bottom: 24px; }
      .paso-num { flex: none; width: 34px; height: 34px; border-radius: 50%; background: #1d1d1f; color: #fff;
        font-weight: 600; font-size: 16px; display: flex; align-items: center; justify-content: center; margin-top: 2px; }
      .paso-title { font-size: 25px; font-weight: 700; letter-spacing: -0.015em; }
      .paso-sub { color: #6e6e73; font-size: 16px; margin-top: 4px; }

      .maestro-card { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 16px; overflow: hidden; }
      .maestro-card.full { border-color: rgba(52,199,89,0.4); }
      .maestro-head { width: 100%; display: flex; align-items: center; gap: 14px; padding: 16px 20px;
        background: none; border: none; cursor: pointer; font: inherit; color: inherit; text-align: left; }
      .maestro-ic { width: 38px; height: 38px; border-radius: 10px; background: #f5f5f7; color: #1d1d1f;
        display: flex; align-items: center; justify-content: center; flex: none; }
      .maestro-card.full .maestro-ic { background: rgba(52,199,89,0.12); color: #248a3d; }
      .maestro-txt { flex: 1; display: flex; flex-direction: column; gap: 2px; min-width: 0; }
      .maestro-txt strong { font-size: 16px; font-weight: 600; }
      .maestro-txt span { font-size: 14px; color: #6e6e73; }
      .maestro-body { padding: 0 20px 20px 72px; }
      .maestro-body p { font-size: 15px; color: #515154; line-height: 1.5; margin-bottom: 14px; }
      .maestro-err { display: inline-flex; align-items: center; gap: 6px; margin-left: 12px; font-size: 14px; color: #c2271c; }

      .grid-tipos { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; }
      .card-tipo { text-align: left; background: #fff; border: 1px solid rgba(0,0,0,0.07); border-radius: 18px;
        padding: 22px; cursor: pointer; font: inherit; color: inherit;
        transition: transform .25s cubic-bezier(.2,.8,.3,1), box-shadow .25s, border-color .25s; }
      .card-tipo:hover { transform: translateY(-3px); box-shadow: 0 14px 34px rgba(0,0,0,0.08); }
      .card-tipo.activo { border-color: #5b8dee; box-shadow: 0 0 0 1px #5b8dee, 0 14px 34px rgba(91,141,238,0.15); }
      .card-tipo-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }
      .card-icon { width: 44px; height: 44px; border-radius: 12px; background: #f5f5f7; color: #1d1d1f;
        display: flex; align-items: center; justify-content: center; }
      .card-icon.sm { width: 36px; height: 36px; border-radius: 10px; flex: none; }
      .card-tipo.activo .card-icon { background: rgba(91,141,238,0.1); color: #5b8dee; }
      .card-check { color: #5b8dee; height: 22px; width: 22px; }
      .card-tipo h3 { font-size: 17px; font-weight: 600; letter-spacing: -0.01em; margin-bottom: 6px; }
      .card-tipo p { font-size: 15px; color: #6e6e73; line-height: 1.45; }

      .zona-cargas { display: flex; flex-direction: column; gap: 30px; }
      .bloque { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 18px; padding: 18px 20px 16px; }
      .bloque-head { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; margin-bottom: 8px; }
      .bloque-tit { flex: 1; min-width: 160px; display: flex; flex-direction: column; gap: 2px; }
      .bloque-tit strong { font-size: 16px; font-weight: 600; }
      .bloque-tit span { font-size: 14px; color: #6e6e73; }
      .bloque-tit em { font-style: normal; font-weight: 600; }
      .bloque-acciones { display: flex; gap: 8px; flex-wrap: wrap; }
      .btn-ghost { display: inline-flex; align-items: center; gap: 6px; background: #f5f5f7; color: #1d1d1f;
        border: none; cursor: pointer; font: inherit; font-size: 14px; font-weight: 500;
        padding: 8px 14px; border-radius: 980px; transition: background .2s; }
      .btn-ghost:hover { background: #e8e8ed; }
      .btn-validar { display: inline-flex; align-items: center; gap: 7px; background: #5b8dee; color: #fff;
        border: none; cursor: pointer; font: inherit; font-size: 14px; font-weight: 500;
        padding: 8px 16px; border-radius: 980px; transition: background .2s, opacity .2s; }
      .btn-validar:hover:not(:disabled) { background: #4a7de0; }
      .btn-validar:disabled { opacity: .4; cursor: not-allowed; }

      .aviso-info { display: flex; align-items: center; gap: 8px; margin: 6px 0 8px; font-size: 14px;
        color: #6e4c00; background: rgba(255,149,0,0.1); border-radius: 10px; padding: 9px 13px; }

      .grilla-scroll { overflow-x: auto; margin-top: 8px; border: 1px solid rgba(0,0,0,0.06); border-radius: 12px; }
      .grilla { border-collapse: collapse; width: 100%; font-size: 14px; }
      .grilla th { font-size: 11px; text-transform: uppercase; letter-spacing: .05em; color: #86868b;
        font-weight: 600; background: #f8f8fc; border-bottom: 1px solid rgba(200,205,230,0.4);
        padding: 8px 10px; text-align: left; white-space: nowrap; }
      .grilla td { padding: 4px 6px; border-bottom: 1px solid rgba(0,0,0,0.04); }
      .grilla tr:last-child td { border-bottom: none; }
      .th-n, .td-n { width: 34px; text-align: center; }
      .th-x, .td-x { width: 70px; text-align: center; }
      .fila-btns { display: inline-flex; gap: 5px; }
      .dz-clear.dup:hover { background: rgba(91,141,238,0.1); color: #5b8dee; }
      .celda { width: 100%; border: 1px solid transparent; background: transparent; font: inherit;
        font-size: 14px; padding: 7px 8px; border-radius: 8px; color: inherit;
        transition: border-color .15s, background .15s; }
      .celda::placeholder { color: #c7c7cc; }
      .celda:hover { background: #f5f5f7; }
      .celda:focus { outline: none; background: #fff; border-color: #5b8dee; box-shadow: 0 0 0 3px rgba(91,141,238,0.15); }
      .g-err td { background: rgba(255,59,48,0.045); }
      .g-ok .td-n { }
      .g-err-detalle td { background: rgba(255,59,48,0.045); color: #c2271c; font-size: 13px;
        padding: 0 10px 8px; border-bottom: 1px solid rgba(0,0,0,0.04); }
      .ic-dim { color: #d2d2d7; }
      .btn-addrow { display: inline-flex; align-items: center; gap: 6px; margin-top: 10px;
        background: none; border: 1px dashed rgba(0,0,0,0.18); color: #515154; cursor: pointer;
        font: inherit; font-size: 14px; padding: 7px 14px; border-radius: 980px; transition: border-color .2s, color .2s; }
      .btn-addrow:hover { border-color: #5b8dee; color: #5b8dee; }

      .btn-primary { display: inline-flex; align-items: center; gap: 8px; background: #5b8dee; color: #fff;
        border: none; cursor: pointer; font: inherit; font-size: 16px; font-weight: 600; padding: 12px 26px;
        border-radius: 14px; transition: background .2s, transform .15s, box-shadow .2s, opacity .2s;
        box-shadow: 0 4px 14px rgba(91,141,238,0.35); }
      .btn-primary:hover:not(:disabled) { background: #4a7de0; transform: translateY(-1px); box-shadow: 0 6px 20px rgba(91,141,238,0.45); }
      .btn-primary:active:not(:disabled) { transform: translateY(0); }
      .btn-primary:disabled { opacity: .35; cursor: not-allowed; }
      .btn-soft { display: inline-flex; align-items: center; gap: 7px; flex: none; background: rgba(91,141,238,0.1);
        color: #5b8dee; border: none; cursor: pointer; font: inherit; font-size: 14.5px; font-weight: 500;
        padding: 9px 16px; border-radius: 12px; transition: background .2s; }
      .btn-soft:hover { background: rgba(91,141,238,0.18); }

      .tabla-wrap { margin-top: 16px; background: #fff; border: 1px solid rgba(200,205,230,0.5);
        border-radius: 16px; overflow: hidden; box-shadow: 0 2px 12px rgba(100,110,180,0.07); }
      .tabla-head { display: flex; justify-content: space-between; align-items: center; gap: 12px;
        padding: 11px 16px; font-size: 14px; color: #515154; border-bottom: 1px solid rgba(0,0,0,0.06); flex-wrap: wrap; }
      .t-ok { color: #248a3d; } .t-err { color: #c2271c; }
      .link-export { display: inline-flex; align-items: center; gap: 5px; background: none; border: none;
        color: #5b8dee; cursor: pointer; font: inherit; font-size: 13.5px; font-weight: 500; }
      .link-export:disabled { color: #c7c7cc; cursor: not-allowed; }
      .tabla-scroll { overflow-x: auto; max-height: 420px; overflow-y: auto; }
      .tabla-wrap table { border-collapse: collapse; width: 100%; font-size: 14px; }
      .tabla-wrap th, .tabla-wrap td { padding: 8px 12px; text-align: left; white-space: nowrap; }
      .tabla-wrap thead th { position: sticky; top: 0; font-size: 11px; text-transform: uppercase;
        letter-spacing: .05em; color: #86868b; font-weight: 600; background: #f8f8fc;
        border-bottom: 1px solid rgba(200,205,230,0.4); }
      .th-auto { color: #5e5ce6 !important; }
      .tabla-wrap tbody tr { border-bottom: 1px solid rgba(0,0,0,0.04); }
      .tabla-wrap tbody tr:last-child { border-bottom: none; }
      .row-err { background: rgba(255,59,48,0.035); }
      .td-auto { color: #5e5ce6; }
      .td-obs { color: #86868b; max-width: 300px; white-space: normal !important; }
      .row-err .td-obs { color: #c2271c; }
      .ic-ok { color: #34c759; } .ic-err { color: #ff3b30; }

      .barra-envio { display: flex; align-items: center; gap: 18px; flex-wrap: wrap; background: #1d1d1f;
        color: #f5f5f7; border-radius: 20px; padding: 20px 26px; }
      .barra-stats { display: flex; gap: 16px; font-size: 15px; font-weight: 500; }
      .stat { display: inline-flex; align-items: center; gap: 6px; }
      .stat.ok { color: #30d158; } .stat.err { color: #ff6961; }
      .barra-msg { flex: 1; font-size: 15px; color: #a1a1a6; min-width: 200px; }

      .success-wrap { max-width: 580px; margin: 0 auto; padding: 100px 22px; text-align: center; }
      .success-check { color: #34c759; display: flex; justify-content: center; animation: pop .5s cubic-bezier(.2,.9,.3,1.4); }
      .success-title { font-size: 40px; font-weight: 700; letter-spacing: -0.018em; margin: 22px 0 12px; }
      .success-sub { font-size: 17px; color: #515154; line-height: 1.55; margin-bottom: 30px; }
      .success-resumen { background: #fff; border: 1px solid rgba(0,0,0,0.07); border-radius: 16px; padding: 8px 0; margin-bottom: 32px; text-align: left; }
      .success-item { display: flex; align-items: center; gap: 12px; padding: 12px 20px; font-size: 15.5px; border-bottom: 1px solid rgba(0,0,0,0.05); }
      .success-item:last-child { border-bottom: none; }
      .success-item > svg { color: #5b8dee; flex: none; }
      .success-item > span:nth-child(2) { flex: 1; }
      .success-count { color: #86868b; font-size: 14px; }

      /* CVP */
      .cvp-flujos { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-top: 4px; }
      .cvp-flujo { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 16px; padding: 20px; display: flex; flex-direction: column; gap: 8px; }
      .cvp-alta { border-top: 3px solid #34c759; }
      .cvp-mod { border-top: 3px solid #ff9500; }
      .cvp-baja { border-top: 3px solid #ff3b30; }
      .cvp-flujo-ic { color: #515154; }
      .cvp-flujo strong { font-size: 16px; font-weight: 600; }
      .cvp-flujo p { font-size: 14px; color: #6e6e73; line-height: 1.5; }
      .cvp-skus { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 4px; }
      .cvp-sku { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 16px; padding: 20px; }
      .cvp-sku-badge { display: inline-block; font-size: 13px; font-weight: 600; border-radius: 980px; padding: 4px 12px; margin-bottom: 10px; }
      .cvp-parcial .cvp-sku-badge { background: #e8f5e9; color: #2e7d32; }
      .cvp-total .cvp-sku-badge { background: #e3f2fd; color: #1565c0; }
      .cvp-sku-desc { font-size: 14.5px; color: #515154; line-height: 1.5; margin-bottom: 12px; }
      .cvp-sku-lista { list-style: none; display: flex; flex-direction: column; gap: 6px; padding: 0; }
      .cvp-sku-lista li { font-size: 14px; color: #3c3c43; padding-left: 16px; position: relative; line-height: 1.4; }
      .cvp-sku-lista li::before { content: "·"; position: absolute; left: 4px; color: #86868b; }
      .cvp-etapas { display: flex; flex-direction: column; gap: 8px; margin-top: 4px; }
      .cvp-etapa { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 14px; overflow: hidden; }
      .cvp-etapa.open { box-shadow: 0 6px 20px rgba(0,0,0,0.06); }
      .cvp-etapa-btn { width: 100%; display: flex; align-items: center; gap: 14px; padding: 16px 18px; background: none; border: none; cursor: pointer; font: inherit; color: inherit; text-align: left; }
      .cvp-etapa-btn:hover { background: #f5f5f7; }
      .cvp-etapa-num { width: 28px; height: 28px; border-radius: 50%; background: #f0f0f2; display: flex; align-items: center; justify-content: center; font-size: 14px; font-weight: 600; color: #515154; flex-shrink: 0; }
      .cvp-etapa.open .cvp-etapa-num { background: #1d1d1f; color: #fff; }
      .cvp-etapa-head { flex: 1; }
      .cvp-etapa-head strong { display: block; font-size: 15.5px; font-weight: 600; }
      .cvp-etapa-head em { display: block; font-size: 13.5px; color: #6e6e73; font-style: normal; margin-top: 2px; }
      .cvp-etapa-toggle { color: #86868b; }
      .cvp-etapa-body { border-top: 1px solid rgba(0,0,0,0.06); padding: 14px 18px; display: flex; flex-direction: column; gap: 12px; }
      .cvp-paso { display: flex; align-items: flex-start; gap: 12px; }
      .cvp-paso-n { width: 22px; height: 22px; border-radius: 50%; background: #f0f0f2; font-size: 11px; font-weight: 700; color: #515154; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 2px; }
      .cvp-paso-cont { flex: 1; }
      .cvp-paso-cont strong { display: block; font-size: 14.5px; font-weight: 600; }
      .cvp-paso-cont span { display: block; font-size: 13.5px; color: #515154; line-height: 1.45; margin-top: 2px; }
      .cvp-paso-resp { font-size: 11.5px; color: #86868b; display: block; margin-top: 3px; }
      .cvp-decision { font-size: 11px; font-weight: 600; background: #fff8e1; color: #f57c00; border-radius: 980px; padding: 3px 9px; flex-shrink: 0; margin-top: 3px; white-space: nowrap; }
      .cvp-actores { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 4px; }
      .cvp-actor { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 14px; padding: 16px; display: flex; flex-direction: column; gap: 4px; }
      .cvp-actor strong { font-size: 14.5px; font-weight: 600; line-height: 1.3; }
      .cvp-actor-area { font-size: 13px; color: #5b8dee; font-weight: 500; }
      .cvp-actor p { font-size: 13.5px; color: #6e6e73; line-height: 1.4; margin-top: 4px; }
      .cvp-raci-wrap { overflow-x: auto; }
      .cvp-raci { width: 100%; border-collapse: collapse; font-size: 14px; }
      .cvp-raci th { background: #f5f5f7; font-weight: 600; padding: 10px 14px; text-align: left; border-bottom: 1px solid rgba(0,0,0,0.08); white-space: nowrap; }
      .cvp-raci td { padding: 9px 14px; border-bottom: 1px solid rgba(0,0,0,0.05); }
      .cvp-raci tr:last-child td { border-bottom: none; }
      .cvp-raci-r { font-weight: 700; color: #1d1d1f; }
      .cvp-raci-a { font-weight: 700; color: #5b8dee; }
      .cvp-raci-c { color: #515154; }
      .cvp-raci-i { color: #86868b; }
      @media (max-width: 700px) {
        .cvp-flujos { grid-template-columns: 1fr; }
        .cvp-skus { grid-template-columns: 1fr; }
        .cvp-actores { grid-template-columns: 1fr 1fr; }
      }

      /* CVP v2 — multi sub-proceso */
      .cvp2-tabs-wrap { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 16px; margin-bottom: 20px; overflow: hidden; }
      .cvp2-tabs { display: flex; overflow-x: auto; scrollbar-width: none; gap: 0; }
      .cvp2-tabs::-webkit-scrollbar { display: none; }
      .cvp2-tab-item { display: flex; align-items: center; gap: 8px; flex-shrink: 0; padding: 14px 18px; background: none; border: none; border-bottom: 3px solid transparent; cursor: pointer; font: inherit; font-size: 14.5px; font-weight: 500; color: #86868b; text-align: left; transition: color .15s, border-color .15s; white-space: nowrap; }
      .cvp2-tab-item:hover { color: #3c3c43; background: #f9f9f9; }
      .cvp2-tab-active { font-weight: 600; }
      .cvp2-tab-num { width: 20px; height: 20px; border-radius: 50%; background: #f0f0f2; font-size: 10.5px; font-weight: 700; color: #515154; display: flex; align-items: center; justify-content: center; flex-shrink: 0; transition: background .15s, color .15s; }
      .cvp2-tab-name { line-height: 1.3; }
      .cvp2-tabs-select { display: none; width: 100%; padding: 14px 16px; font: inherit; font-size: 15px; font-weight: 500; color: #1d1d1f; background: #fff; border: none; border-radius: 0; appearance: none; -webkit-appearance: none; background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%2386868b' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 16px center; cursor: pointer; }
      .cvp2-main { display: flex; flex-direction: column; gap: 14px; }
      .cvp2-header { border-top: 4px solid #ccc; border-radius: 16px; background: #f5f5f7; padding: 24px 28px; }
      .cvp2-sp-num { font-size: 11px; font-weight: 700; color: #fff; border-radius: 980px; padding: 3px 10px; }
      .cvp2-estado { font-size: 11.5px; font-weight: 600; background: #fff3e0; color: #e65100; border-radius: 980px; padding: 3px 10px; }
      .cvp2-version { font-size: 11.5px; color: #86868b; font-weight: 500; }
      .cvp2-title { font-size: 22px; font-weight: 700; letter-spacing: -0.015em; color: #1d1d1f; margin: 0; }
      .cvp2-sec { background: #fff; border: 1px solid rgba(200,205,230,0.5); border-radius: 16px; overflow: hidden;
        box-shadow: 0 2px 10px rgba(100,110,180,0.06); }
      .cvp2-sec-btn { width: 100%; display: flex; align-items: center; justify-content: space-between; padding: 18px 24px; background: none; border: none; cursor: pointer; font: inherit; font-size: 16px; font-weight: 600; color: #1d1d1f; text-align: left; }
      .cvp2-sec-btn:hover { background: #f9f9f9; }
      .cvp2-sec-body { border-top: 1px solid rgba(0,0,0,0.06); padding: 22px 24px; }
      .cvp2-proposito { font-size: 15px; color: #3c3c43; line-height: 1.65; margin: 0 0 18px; }
      .cvp2-alcance-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
      .cvp2-alcance-item { font-size: 14px; color: #515154; line-height: 1.55; background: #f5f5f7; border-radius: 12px; padding: 14px 16px; }
      .cvp2-alcance-item strong { display: block; color: #1d1d1f; font-size: 13px; margin-bottom: 4px; }
      .cvp2-actores-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
      .cvp2-actor-card { background: #f9f9f9; border: 1px solid rgba(0,0,0,0.06); border-radius: 14px; padding: 16px 18px; }
      .cvp2-actor-rol { font-size: 15px; font-weight: 600; color: #1d1d1f; margin-bottom: 3px; line-height: 1.3; }
      .cvp2-actor-area { font-size: 13px; font-weight: 500; margin-bottom: 8px; }
      .cvp2-actor-resp { font-size: 13.5px; color: #515154; line-height: 1.5; }
      .cvp2-raci-scroll { overflow-x: auto; }
      .cvp2-raci-tbl { width: 100%; border-collapse: collapse; font-size: 14px; }
      .cvp2-raci-tbl th { background: #f5f5f7; font-weight: 600; padding: 11px 14px; text-align: left; border-bottom: 1px solid rgba(0,0,0,0.08); white-space: nowrap; font-size: 13.5px; }
      .cvp2-raci-tbl td { padding: 10px 14px; border-bottom: 1px solid rgba(0,0,0,0.05); text-align: center; font-weight: 700; font-size: 14px; min-width: 50px; }
      .cvp2-raci-tbl td:first-child { text-align: left; font-weight: 400; color: #3c3c43; min-width: 220px; }
      .cvp2-raci-tbl tr:last-child td { border-bottom: none; }
      .cvp2-r { color: #1d1d1f !important; background: #f0f0f2; }
      .cvp2-a { color: #5b8dee !important; }
      .cvp2-c { color: #515154 !important; }
      .cvp2-i { color: #aeaeb2 !important; }
      .cvp2-raci-legend { font-size: 13px; color: #86868b; margin-top: 12px; }
      .cvp2-fase { margin-bottom: 28px; }
      .cvp2-fase:last-child { margin-bottom: 0; }
      .cvp2-fase-title { font-size: 13.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #515154; border-left: 3px solid; padding-left: 12px; margin-bottom: 14px; }
      .cvp2-pasos-list { display: flex; flex-direction: column; gap: 10px; }
      .cvp2-paso { display: flex; align-items: flex-start; gap: 12px; padding: 14px 16px; background: #f9f9f9; border-radius: 12px; border: 1px solid rgba(0,0,0,0.05); }
      .cvp2-paso-decision { border-style: dashed; }
      .cvp2-paso-n { width: 26px; height: 26px; border-radius: 50%; background: #e5e5ea; font-size: 11.5px; font-weight: 700; color: #515154; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 1px; transition: background .15s; font-variant-numeric: tabular-nums; }
      .cvp2-paso-cont { flex: 1; }
      .cvp2-paso-cont strong { display: block; font-size: 14.5px; font-weight: 600; color: #1d1d1f; margin-bottom: 3px; }
      .cvp2-paso-cont span { display: block; font-size: 13.5px; color: #515154; line-height: 1.5; margin-bottom: 4px; }
      .cvp2-paso-cont em { font-size: 11.5px; color: #86868b; font-style: normal; }
      .cvp2-decision-tag { font-size: 11px; font-weight: 600; border-radius: 980px; padding: 3px 9px; flex-shrink: 0; white-space: nowrap; margin-top: 2px; }
      @media (max-width: 900px) {
        .cvp2-alcance-grid { grid-template-columns: 1fr; }
        .cvp2-actores-grid { grid-template-columns: 1fr 1fr; }
      }
      @media (max-width: 700px) {
        .cvp2-tabs { display: none; }
        .cvp2-tabs-select { display: block; }
        .cvp2-tabs-wrap { border-radius: 12px; }
        .cvp2-header { padding: 18px 16px; }
        .cvp2-title { font-size: 18px; }
        .cvp2-sec-btn { padding: 14px 16px; font-size: 15px; }
        .cvp2-sec-body { padding: 16px; }
        .cvp2-actores-grid { grid-template-columns: 1fr; }
        .cvp2-alcance-grid { grid-template-columns: 1fr; }
      }

      /* ── VISTA INICIO ── */
      .inicio-wrap { padding-top: 0 !important; }
      .inicio-hero { display: grid; grid-template-columns: 1fr 1fr; border-radius: 20px; overflow: hidden; background: #0c0c14; margin-bottom: 28px; min-height: 280px; }
      .inicio-hero-visual { position: relative; overflow: hidden; background: radial-gradient(ellipse at 60% 80%, #1a0800 0%, #0c0c14 70%); }
      .inicio-planet-fire { position: absolute; width: 180px; height: 180px; border-radius: 50%; background: radial-gradient(circle at 38% 38%, #ff6b2b 0%, #c0392b 40%, #7f0000 70%, #1a0000 100%); bottom: 14%; left: 50%; transform: translateX(-50%); box-shadow: 0 0 80px rgba(220,60,0,0.35), inset -12px -8px 30px rgba(0,0,0,0.6); }
      .inicio-planet-earth { position: absolute; width: 300px; height: 300px; border-radius: 50%; background: radial-gradient(circle at 32% 28%, #1a3a5c 0%, #0d2137 55%, #04101e 100%); top: -90px; right: -70px; box-shadow: 0 0 100px rgba(20,60,130,0.25), inset -20px -12px 50px rgba(0,0,0,0.5); }
      .inicio-stars { position: absolute; inset: 0; background-image: radial-gradient(circle, rgba(255,255,255,0.7) 1px, transparent 1px); background-size: 40px 40px; opacity: 0.18; }
      .inicio-hero-text { padding: 28px 36px; display: flex; flex-direction: column; justify-content: center; gap: 16px; }
      .inicio-logomark { font-size: 18px; font-weight: 800; color: rgba(255,255,255,0.35); letter-spacing: 0.1em; }
      .inicio-title { font-size: 21px; font-weight: 700; color: #ffffff; line-height: 1.45; letter-spacing: -0.01em; margin: 0; }
      .inicio-desc-blue { font-size: 15px; color: #6fa8f5; line-height: 1.7; margin: 0; font-weight: 500; }
      .inicio-desc-orange { font-size: 15px; color: #ffb347; line-height: 1.7; margin: 0; font-weight: 500; }
      .inicio-logos { display: flex; align-items: center; justify-content: space-around; flex-wrap: wrap; gap: 24px; margin: 20px 0 4px; }
      .inicio-logo { height: 52px; width: auto; object-fit: contain; opacity: 0.85; transition: opacity .2s; }
      .inicio-logo:hover { opacity: 1; }
      .inicio-section-title { font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; color: #86868b; margin-bottom: 12px; }
      .inicio-modules { display: flex; flex-direction: column; gap: 10px; }
      .inicio-card { display: flex; align-items: center; gap: 16px; padding: 18px 20px; background: #fff; border: 1px solid rgba(0,0,0,0.07); border-radius: 16px; cursor: pointer; font: inherit; text-align: left; transition: all .18s; box-shadow: 0 2px 8px rgba(0,0,0,0.04); width: 100%; }
      .inicio-card:hover { transform: translateX(4px); box-shadow: 0 4px 20px rgba(0,0,0,0.09); border-color: rgba(0,0,0,0.13); }
      .inicio-card-icon { width: 44px; height: 44px; border-radius: 12px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
      .inicio-card-body { flex: 1; }
      .inicio-card-body strong { display: block; font-size: 15.5px; font-weight: 600; color: #1d1d1f; margin-bottom: 3px; }
      .inicio-card-body span { font-size: 14px; color: #86868b; line-height: 1.4; }
      .inicio-card-arrow { color: #c7c7cc; flex-shrink: 0; }
      @media (max-width: 700px) {
        .inicio-hero { grid-template-columns: 1fr; }
        .inicio-hero-visual { min-height: 180px; }
        .inicio-hero-text { padding: 28px 22px; gap: 16px; }
        .inicio-title { font-size: 18px; }
      }

      /* ADMIN */
      /* ── Mantenedor de bases maestras ── */
      .mto-tabs { display: flex; gap: 8px; margin-bottom: 20px; flex-wrap: wrap; }
      .mto-tab { display: flex; align-items: center; gap: 7px; padding: 9px 18px; border-radius: 980px; border: 1.5px solid rgba(0,0,0,0.1); background: #fff; font: inherit; font-size: 15px; font-weight: 500; color: #515154; cursor: pointer; transition: all .18s; }
      .mto-tab:hover { border-color: rgba(0,0,0,0.22); color: #1d1d1f; }
      .mto-tab.on { background: #1d1d1f; color: #fff; border-color: #1d1d1f; }
      .mto-badge { margin-left: 4px; font-size: 11px; font-weight: 700; padding: 2px 7px; border-radius: 980px; }
      .mto-badge.ok { background: rgba(52,199,89,0.15); color: #248a3d; }
      .mto-badge.err { background: rgba(255,59,48,0.12); color: #c2271c; }
      .mto-panel { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 18px; padding: 28px; }
      .mto-desc { font-size: 15px; color: #515154; margin: 0 0 20px; }
      .mto-actions { display: flex; gap: 10px; flex-wrap: wrap; }
      .mto-msg { display: flex; align-items: center; gap: 8px; font-size: 15px; padding: 12px 16px; border-radius: 10px; margin-top: 18px; font-weight: 500; }
      .mto-msg.ok { background: rgba(52,199,89,0.1); color: #248a3d; }
      .mto-msg.err { background: rgba(255,59,48,0.08); color: #c2271c; }
      .mto-msg.error { background: rgba(255,59,48,0.08); color: #c2271c; }
      .mto-summary { display: flex; align-items: center; gap: 14px; margin: 18px 0 12px; flex-wrap: wrap; }
      .mto-file { font-size: 14px; color: #1d1d1f; font-weight: 600; }
      .mto-cnt { display: flex; align-items: center; gap: 5px; font-size: 14px; font-weight: 600; padding: 4px 10px; border-radius: 980px; }
      .mto-cnt.ok { background: rgba(52,199,89,0.1); color: #248a3d; }
      .mto-cnt.err { background: rgba(255,59,48,0.08); color: #c2271c; }
      .mto-table-wrap { overflow-x: auto; border: 1px solid rgba(0,0,0,0.08); border-radius: 12px; margin-top: 4px; }
      .mto-table { width: 100%; border-collapse: collapse; font-size: 14px; }
      .mto-table th { background: #f5f5f7; font-weight: 600; padding: 10px 14px; text-align: left; border-bottom: 1px solid rgba(0,0,0,0.08); white-space: nowrap; }
      .mto-table td { padding: 9px 14px; border-bottom: 1px solid rgba(0,0,0,0.04); white-space: nowrap; }
      .mto-table tr:last-child td { border-bottom: none; }
      .mto-table tr.fila-err td { background: rgba(255,59,48,0.03); }
      .mto-vacio { color: #b0b0b5; font-style: italic; }
      .mto-mas { font-size: 13px; color: #86868b; text-align: center; padding: 10px; }
      /* ── Admin ── */
      .admin-form { background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 16px; padding: 20px; }
      .admin-form-row { display: grid; grid-template-columns: 1fr 1fr 1fr auto; gap: 14px; align-items: start; }
      .admin-select { border: 1px solid rgba(0,0,0,0.15); border-radius: 10px; padding: 10px 14px; font: inherit; font-size: 16px; outline: none; width: 100%; background: #fff; color: #1d1d1f; }
      .admin-msj-ok { font-size: 14px; color: #34c759; display: flex; align-items: center; gap: 6px; }
      .admin-tabla-wrap { overflow-x: auto; background: #fff; border: 1px solid rgba(0,0,0,0.08); border-radius: 16px; }
      .admin-tabla { width: 100%; border-collapse: collapse; font-size: 15px; }
      .admin-tabla th { background: #f5f5f7; font-weight: 600; padding: 12px 16px; text-align: left; border-bottom: 1px solid rgba(0,0,0,0.08); font-size: 14px; }
      .admin-tabla td { padding: 12px 16px; border-bottom: 1px solid rgba(0,0,0,0.05); }
      .admin-tabla tr:last-child td { border-bottom: none; }
      .admin-rol-sel { border: 1px solid rgba(0,0,0,0.12); border-radius: 980px; padding: 4px 12px; font: inherit; font-size: 13px; font-weight: 600; cursor: pointer; outline: none; }
      .rol-ddmm { background: #e3f2fd; color: #1565c0; border-color: #bbdefb; }
      .rol-sol { background: #f5f5f7; color: #515154; }
      .admin-activo { border: none; border-radius: 980px; padding: 5px 14px; font: inherit; font-size: 13px; font-weight: 600; cursor: pointer; }
      .activo-on { background: #e8f5e9; color: #2e7d32; }
      .activo-off { background: #fff3e0; color: #e65100; }
      @media (max-width: 700px) { .admin-form-row { grid-template-columns: 1fr; } }

      /* LOGIN */
      .login-wrap { min-height: 100vh; display: flex; align-items: center; justify-content: center; background: #e4e7f1; padding: 20px; }
      .login-box { background: #fff; border: 1px solid rgba(200,205,230,0.5); border-radius: 20px; padding: 40px 36px; width: 100%; max-width: 400px; box-shadow: 0 8px 40px rgba(100,110,180,0.14); }
      .login-logo { font-weight: 700; font-size: 16px; color: #1d1d1f; margin-bottom: 24px; }
      .login-title { font-size: 28px; font-weight: 700; letter-spacing: -0.015em; margin-bottom: 6px; }
      .login-sub { font-size: 14.5px; color: #6e6e73; margin-bottom: 28px; line-height: 1.4; }
      .login-form { display: flex; flex-direction: column; gap: 16px; }
      .login-field { display: flex; flex-direction: column; gap: 6px; }
      .login-field label { font-size: 14px; font-weight: 500; color: #3c3c43; }
      .login-field input { border: 1px solid rgba(0,0,0,0.15); border-radius: 10px; padding: 10px 14px; font: inherit; font-size: 16px; outline: none; transition: border-color .2s; color: #1d1d1f; }
      .login-field input:focus { border-color: #5b8dee; box-shadow: 0 0 0 3px rgba(91,141,238,0.15); }
      .login-error { font-size: 14px; color: #ff3b30; display: flex; align-items: center; gap: 6px; }
      .login-loading { display: flex; align-items: center; justify-content: center; min-height: 100vh; font-size: 16px; color: #86868b; background: #e4e7f1; }
      .nav-nombre { font-size: 13px; color: #6e6e73; }
      .nav-logout { border: 1px solid rgba(0,0,0,0.1); background: none; cursor: pointer; font: inherit; font-size: 13px; font-weight: 500; color: #515154; padding: 5px 12px; border-radius: 980px; transition: all .2s; }
      .nav-logout:hover { background: rgba(255,59,48,0.07); color: #ff3b30; border-color: rgba(255,59,48,0.2); }
      .nav-hamburger { display: none; border: none; background: none; cursor: pointer; color: #3c3c43; padding: 4px; border-radius: 8px; transition: background .15s; }
      .nav-hamburger:hover { background: rgba(0,0,0,0.06); }

      /* Mobile overlay menu */
      .nav-mobile-overlay { position: fixed; inset: 0; z-index: 200; background: rgba(20,22,42,0.35);
        backdrop-filter: blur(4px); -webkit-backdrop-filter: blur(4px); }
      .nav-mobile-menu { position: absolute; top: 52px; left: 0; right: 0; background: #fff;
        border-bottom: 1px solid rgba(200,205,230,0.5); padding: 8px 0 16px;
        box-shadow: 0 12px 40px rgba(100,110,180,0.16); animation: slideDown .2s cubic-bezier(.2,.8,.3,1); }
      @keyframes slideDown { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: none; } }
      .nav-mobile-user { font-size: 13px; font-weight: 600; color: #86868b; padding: 10px 20px 8px;
        text-transform: uppercase; letter-spacing: .06em; border-bottom: 1px solid rgba(200,205,230,0.4); margin-bottom: 6px; }
      .nav-mobile-item { display: block; width: 100%; text-align: left; background: none; border: none;
        font: inherit; font-size: 16px; font-weight: 500; color: #1d1d1f; padding: 13px 20px;
        cursor: pointer; transition: background .15s; }
      .nav-mobile-item:hover { background: rgba(228,231,241,0.5); }
      .nav-mobile-item.on { color: #5b8dee; font-weight: 600; background: rgba(91,141,238,0.06); }
      .nav-mobile-salir { display: block; width: 100%; text-align: left; background: none; border: none;
        font: inherit; font-size: 16px; color: #ff3b30; padding: 13px 20px; cursor: pointer;
        border-top: 1px solid rgba(200,205,230,0.4); margin-top: 6px; transition: background .15s; }
      .nav-mobile-salir:hover { background: rgba(255,59,48,0.05); }
      .login-link { background: none; border: none; cursor: pointer; font: inherit; font-size: 14px; color: #5b8dee; padding: 0; margin-top: 16px; display: block; text-align: center; width: 100%; }
      .login-link:hover { text-decoration: underline; }
      .login-ok { font-size: 15px; color: #34c759; display: flex; align-items: center; gap: 8px; margin-top: 8px; }

      .pie { text-align: center; color: #86868b; font-size: 13.5px; padding: 30px 22px 50px; }
      .dz-clear { border: none; cursor: pointer; width: 26px; height: 26px; border-radius: 50%;
        background: #f5f5f7; color: #6e6e73; display: inline-flex; align-items: center;
        justify-content: center; transition: background .2s, color .2s; }
      .dz-clear:hover { background: rgba(255,59,48,0.1); color: #ff3b30; }
      .fade-in { animation: fade .45s cubic-bezier(.2,.8,.3,1); }
      @keyframes fade { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: none; } }
      @keyframes pop { from { opacity: 0; transform: scale(.6); } to { opacity: 1; transform: scale(1); } }
      .card-tipo:focus-visible, button:focus-visible { outline: 2px solid #5b8dee; outline-offset: 2px; }
      @media (prefers-reduced-motion: reduce) { .fade-in, .success-check { animation: none; } .card-tipo, .btn-primary { transition: none; } }
      @media (max-width: 640px) {
        .barra-envio { flex-direction: column; align-items: stretch; text-align: center; }
        .barra-stats { justify-content: center; }
        .maestro-body { padding-left: 20px; }
        .bloque-acciones { width: 100%; }

        /* Títulos y espaciados */
        .sol-wrap { padding: 24px 14px 48px; }
        .sol-title { font-size: 26px; }
        .sol-sub { font-size: 14.5px; }
        .seccion { padding: 14px 14px 24px; }
        .hero { padding: 48px 14px 32px; }
        .hero-sub { font-size: 16px; }

        /* Nav mobile */
        .nav-tabs-desktop { display: none; }
        .nav-hamburger { display: flex; }
        .nav-badge { display: none; }
        .nav-logout { display: none; }
        .nav-nombre { display: none; }

        /* Cards y grillas */
        .grid-tipos { grid-template-columns: 1fr; }
        .clu-grid { grid-template-columns: 1fr; }
        .ayuda-validaciones { grid-template-columns: 1fr; }
        .cvp-actores { grid-template-columns: 1fr 1fr; }
        .cvp-flujos { grid-template-columns: 1fr; }
        .cvp-skus { grid-template-columns: 1fr; }
        .cvp2-alcance-grid { grid-template-columns: 1fr; }
        .cvp2-actores-grid { grid-template-columns: 1fr; }

        /* Solicitudes */
        .sol-head { flex-direction: column; gap: 10px; }
        .sol-fila { flex-wrap: wrap; gap: 8px; }
        .sol-folio { min-width: unset; }
        .sol-card { border-radius: 14px; }

        /* Admin form */
        .admin-form-row { grid-template-columns: 1fr; }
        .login-box { padding: 28px 20px; }

        /* Bloque de planillas */
        .bloque-head { flex-direction: column; align-items: flex-start; }
        .paso-head { flex-direction: column; gap: 8px; }

        /* Shell mobile: sidebar oculta, topbar visible */
        .shell { flex-direction: column; }
        .shell-body { flex-direction: column; }
        .shell-content { padding: 16px 14px; }
        .shell-right { width: 100%; height: auto; position: static; border-left: none; border-top: 1px solid rgba(0,0,0,0.07); }
        .sidebar { display: none; }
        .topbar-mobile { display: flex; align-items: center; justify-content: space-between;
          padding: 0 16px; height: 52px; background: #f2f2f7; color: #1d1d1f;
          position: sticky; top: 0; z-index: 50; border-bottom: 1px solid rgba(0,0,0,0.1); }
      }
    `}</style>
  );
}
