import streamlit as st
import streamlit.components.v1 as components
import json
import random
from datetime import datetime
from checker import revisar_premios
from notificar import enviar_correo_con_pdf
from scraper import (
    obtener_resultados, 
    obtener_resultados_por_numero, 
    obtener_fecha_sorteo_actual,
    obtener_sorteos_guardados,
    importar_desde_excel,
    exportar_historial_a_excel
)

st.set_page_config(page_title="🎰 Quini 6 Checker", page_icon="🎯", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stButton > button { width: 100%; }
    .numero-bola { display: inline-block; width: 40px; height: 40px; line-height: 40px; border-radius: 50%; background: #FFD700; color: #000; font-weight: bold; margin: 3px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }
    iframe { border: none !important; }
    .badge-vacante { display: inline-block; padding: 4px 12px; border-radius: 12px; background: #e74c3c; color: white; font-weight: bold; font-size: 0.8em; margin-top: 5px; }
    .badge-ganado { display: inline-block; padding: 4px 12px; border-radius: 12px; background: #27ae60; color: white; font-weight: bold; font-size: 0.8em; margin-top: 5px; }
    .badge-desconocido { display: inline-block; padding: 4px 12px; border-radius: 12px; background: #95a5a6; color: white; font-weight: bold; font-size: 0.8em; margin-top: 5px; }
    @media (max-width: 768px) { .stColumns { gap: 5px !important; } .numero-bola { width: 30px; height: 30px; line-height: 30px; font-size: 0.8em; margin: 1px; } }
</style>
""", unsafe_allow_html=True)

st.title("🎰 Quini 6 Checker")
st.markdown("### Resultados y control de jugadas")

for key, val in [("ultimo_chequeo", None), ("resultados_cache", None), ("pozos_cache", None), ("info_sorteo_cache", None), ("sorteo_seleccionado", "ultimo"), ("mostrar_detalle", False)]:
    if key not in st.session_state: st.session_state[key] = val

@st.cache_data(ttl=60)
def cargar_jugadas():
    try:
        with open("jugadas.json", "r", encoding="utf-8") as f: return json.load(f)
    except:
        return [
            {"nombre": "Adrian", "email": "", "numeros": [7, 15, 18, 23, 33, 34]},
            {"nombre": "Carlos", "email": "", "numeros": [5, 6, 22, 24, 39, 40]},
            {"nombre": "Maxi", "email": "", "numeros": [2, 9, 10, 14, 19, 41]},
            {"nombre": "Ruben", "email": "", "numeros": [3, 7, 17, 21, 33, 36]},
        ]

def guardar_jugadas(j): 
    with open("jugadas.json", "w", encoding="utf-8") as f: json.dump(j, f, indent=2, ensure_ascii=False)
    st.cache_data.clear()

def _obtener_mensaje_aliento(nombre, aciertos_totales, mejor_modalidad, mejor_aciertos, numero_sorteo=""):
    if mejor_aciertos >= 6:
        mensajes = [
            f"🏆 ¡{nombre}, HISTÓRICO! 6 aciertos en {mejor_modalidad}. El Quini se rinde a tus pies. ¡Que empiece la fiesta! 🍾",
            f"👑 {nombre}, la realeza del Quini tiene nuevo monarca. 6 en {mejor_modalidad}. ¡Inclinémonos ante Su Majestad!",
            f"🎰 ¡JACKPOT, JACKPOT, JACKPOT! {nombre} la rompió toda con 6 en {mejor_modalidad}. ¿Vamos a medias? 😜",
            f"💎 {nombre}, ¡SEIS ACIERTOS en {mejor_modalidad}! Ni Nostradamus lo vio venir. A disfrutar la fortuna.",
            f"🦄 {nombre} encontró el unicornio del Quini: 6 en {mejor_modalidad}. Algo imposible hecho realidad.",
        ]
    elif mejor_aciertos >= 5:
        mensajes = [
            f"🔥 {nombre}, ¡5 aciertos en {mejor_modalidad}! Te quedaste con la miel en los labios. La próxima es tuya.",
            f"🎯 ¡Tiro de francotirador, {nombre}! 5 en {mejor_modalidad}. El premio gordo te hizo ojitos.",
            f"💫 {nombre} estuvo a UN numerito de la gloria en {mejor_modalidad}. 5 aciertos que saben a poco... ¡pero alimentan!",
            f"🚀 {nombre} despegó con 5 en {mejor_modalidad}. ¿La próxima? ¡Órbita de premio mayor!",
            f"🎪 {nombre}, 5 aciertos en {mejor_modalidad}. El circo del Quini casi te contrata como estrella principal.",
        ]
    elif mejor_aciertos >= 4:
        mensajes = [
            f"👀 {nombre} metió 4 en {mejor_modalidad}. Algo de ruido hiciste... ¡la billetera va a sonar! 💵",
            f"🎉 ¡4 aciertos en {mejor_modalidad}, {nombre}! No es para tirar manteca al techo, pero algo vas a cobrar.",
            f"🍀 {nombre}, 4 en {mejor_modalidad}. La diosa fortuna te guiñó un ojo. No es romance todavía, pero coquetea.",
            f"🤑 {nombre}, ¡4 aciertos! La caja chica del Quini tiene tu nombre anotado. Pasá a cobrar.",
            f"🎊 {nombre} hizo 4 en {mejor_modalidad}. El Quini te dice 'te quiero, pero no tanto'. Al menos no es zona de amigos.",
        ]
    elif mejor_aciertos >= 3:
        mensajes = [
            f"🤏 ¡Por un pelín, {nombre}! 3 en {mejor_modalidad}. Como dijo el peluquero: cortita pero al pie.",
            f"🎲 {nombre} y sus 3 aciertos en {mejor_modalidad}. El Quini te está midiendo... y le gustás.",
            f"🔮 {nombre}, 3 en {mejor_modalidad}. Las cartas del tarot dicen que tu momento se acerca. Paciencia, joven padawan.",
            f"🧘 {nombre} alcanzó el nirvana de los 3 aciertos en {mejor_modalidad}. Ni frío ni calor... tibio, pero con estilo.",
            f"🪷 {nombre}, 3 en {mejor_modalidad}. El universo conspira... pero hoy se quedó dormido. Mañana será otro día.",
        ]
    elif mejor_aciertos >= 2:
        mensajes = [
            f"🙂 {nombre} rescató 2 en {mejor_modalidad}. No es mucho, pero es trabajo honesto. ¡Como el sueldo mínimo!",
            f"🎈 {nombre}, 2 aciertos. El Quini te dijo 'hola' de lejos. Al menos no te ignoró completamente.",
            f"🫤 {nombre} y sus 2 aciertos en {mejor_modalidad}. Es como ir a una cita y que solo te den la mano. Frío, muy frío.",
            f"🐢 {nombre}, 2 en {mejor_modalidad}. Lento pero seguro... bueno, lento seguro. La próxima acelerás.",
            f"🌱 {nombre} plantó 2 semillitas en {mejor_modalidad}. Con riego y paciencia, quizás crezca algo. O no.",
        ]
    elif mejor_aciertos >= 1:
        mensajes = [
            f"😅 {nombre} y su solitario acierto en {mejor_modalidad}. Es como tener un solo pelo en la cabeza: técnicamente hay, pero...",
            f"🤷 {nombre}, 1 acierto. El Quini te dio un caramelo para que no llores. La próxima trae la bolsa entera.",
            f"🫣 {nombre} metió 1 en {mejor_modalidad}. Estás mirando la fiesta desde la ventana. ¡Entrá que hay calor!",
            f"🕯️ {nombre}, 1 acierto. Una velita en la oscuridad del azar. Al menos no es apagón total.",
            f"📡 {nombre} captó una señal débil en {mejor_modalidad}: 1 acierto. La antena funciona, pero le falta potencia.",
        ]
    else:
        mensajes = [
            f"😂 {nombre} y su cosecha de 0 aciertos. Ni Mandrake podría haber hecho menos. ¡A otra cosa, mariposa!",
            f"🤡 {nombre}, 0 aciertos. El Quini te aplicó la ley del hielo. Frío polar en {mejor_modalidad}.",
            f"👻 {nombre} jugó de fantasma en este sorteo: 0 aciertos. Ni se notó que participaste.",
            f"🎭 {nombre} protagonizó la tragedia del 0 en {mejor_modalidad}. Pero como dijo el poeta: 'lo importante es participar'... MENTIRA, lo importante es ganar.",
            f"🦗 {nombre}, 0 aciertos. El Quini te respondió con un grillo. Silencio absoluto. La revancha es obligatoria.",
            f"🍩 {nombre} se fue con un rosquete de 0 en {mejor_modalidad}. Cero en orto, como diría un astrólogo optimista.",
            f"🧊 {nombre}, 0 aciertos. Más frío que un pingüino en un freezer. La suerte se tomó vacaciones.",
            f"📉 {nombre} cotizó en baja: 0 aciertos. Las acciones del Quini están en rojo. Que no cunda el pánico.",
        ]
    
    semilla = sum(ord(c) for c in f"{nombre}{numero_sorteo}")
    rng = random.Random(semilla)
    return rng.choice(mensajes)

def _construir_tablero_html(detalle, resultados):
    modalidades = ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]
    html = """<html><head><meta charset="UTF-8"><style>
*{box-sizing:border-box}body{margin:0;padding:10px;background:#fff;font-family:sans-serif}
.tabla-jugador{width:100%;border-collapse:collapse;margin-bottom:20px;border:2px solid #1a3a5c}
.tabla-jugador th.numero{background:#1a3a5c;color:#fff;padding:14px 6px;font-size:1.15em;font-weight:bold;min-width:48px}
.tabla-jugador th.nombre{background:#1a3a5c;color:#fff;padding:14px 15px;font-size:1.2em;font-weight:bold;letter-spacing:1px}
.tabla-jugador th.aciertos-header{background:#1a3a5c;color:#fff;padding:14px 10px;font-size:1em;font-weight:bold}
.tabla-jugador td{padding:11px 6px;font-size:1.05em;color:#000;font-weight:600;background:#fff;min-width:48px;border:1px solid #d5d8dc}
.tabla-jugador td.modalidad{text-align:left;padding-left:15px;font-weight:bold;color:#1a3a5c;background:#eaf0f8;border-right:3px solid #1a3a5c}
.celda-ok{background:#1b7a3d!important;color:#fff!important;font-weight:bold}
.celda-vacia{background:#fff!important;color:#ccc}
.total-premio{background:#1b7a3d!important;color:#fff!important;font-weight:bold;font-size:1.2em!important}
.total-medio{background:#e67e22!important;color:#fff!important;font-weight:bold;font-size:1.1em!important}
.total-bajo{background:#fff!important;color:#333!important;font-weight:bold}
.total-cero{background:#f5f5f5!important;color:#ccc!important}
</style></head><body>"""
    for jd in detalle:
        nombre = jd["nombre"]
        nums_j = list(jd["modalidades"].values())[0]["numeros_jugados"]
        html += f'<table class="tabla-jugador"><thead><tr><th class="nombre">{nombre}</th>'
        for n in nums_j: html += f'<th class="numero">{str(n).zfill(2)}</th>'
        html += '<th class="aciertos-header">Aciertos</th></tr></thead><tbody>'
        for mod in modalidades:
            if mod in jd["modalidades"]:
                d = jd["modalidades"][mod]; a = d["aciertos"]
                if mod in ["Premio Extra", "Revancha"]:
                    ct = "total-premio" if a>=6 else ("total-medio" if a>=3 else ("total-bajo" if a>0 else "total-cero"))
                else:
                    ct = "total-premio" if a>=4 else ("total-medio" if a>=3 else ("total-bajo" if a>0 else "total-cero"))
                nm = "Extra" if mod=="Premio Extra" else mod
                html += f'<tr><td class="modalidad">{nm}</td>'
                for n in nums_j: html += '<td class="celda-ok">OK</td>' if n in d["numeros_sorteados"] else '<td class="celda-vacia">-</td>'
                html += f'<td class="{ct}">{a}</td></tr>'
        html += '</tbody></table>'
    return html + '</body></html>'

def _construir_html_completo_con_mensajes(resultados, pozos, info, detalle, mensajes_aliento):
    css = """<style>
body{font-family:Arial,sans-serif;background:#fff;color:#333;margin:0;padding:20px}
h1{color:#1a3a5c;text-align:center;font-size:2em}
h2{color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:5px;font-size:1.4em;margin-top:25px}
h3{color:#2c5aa0;font-size:1.2em}
.contenedor-modalidades{display:flex;flex-wrap:wrap;gap:10px;margin:15px 0}
.tarjeta-modalidad{flex:1;min-width:22%;background:linear-gradient(135deg,#1a3a5c,#2c5aa0);color:#fff;padding:18px;border-radius:12px;text-align:center}
.tarjeta-modalidad h4{font-size:1.2em;text-transform:uppercase;margin:0 0 10px;color:#fff}
.bolita{display:inline-block;width:35px;height:35px;line-height:35px;border-radius:50%;background:#FFD700;color:#000;font-weight:bold;margin:2px;font-size:.9em}
.bolita-extra{display:inline-block;width:28px;height:28px;line-height:28px;border-radius:50%;background:#FFD700;color:#000;font-weight:bold;margin:1px;font-size:.7em}
.pozo-monto{font-size:1.1em;margin-top:10px;font-weight:bold;color:#FFD700;background:rgba(0,0,0,.3);padding:5px 10px;border-radius:8px;display:inline-block}
.info-ganadores{font-size:.85em;margin-top:6px;color:#fff}
.badge-vacante{display:inline-block;padding:5px 14px;border-radius:10px;background:#e74c3c;color:#fff;font-weight:bold;font-size:.9em;margin-top:8px}
.badge-ganado{display:inline-block;padding:5px 14px;border-radius:10px;background:#27ae60;color:#fff;font-weight:bold;font-size:.9em;margin-top:8px}
.tarjeta-extra{background:linear-gradient(135deg,#6d1a8a,#9b59b6);color:#fff;padding:18px;border-radius:12px;text-align:center;margin:15px 0}
.tabla-pdf{width:100%;border-collapse:collapse;margin:10px 0;border:2px solid #1a3a5c}
.tabla-pdf th{background:#1a3a5c;color:#fff;padding:10px 8px;font-weight:bold}
.tabla-pdf td{padding:10px 8px;background:#fff;color:#000;font-weight:600;border:1px solid #d5d8dc}
.tabla-pdf td.mod{text-align:left;color:#1a3a5c;background:#eaf0f8;font-weight:bold}
.ok-pdf{background:#1b7a3d!important;color:#fff!important;font-weight:bold}
.total-verde{background:#1b7a3d!important;color:#fff!important;font-weight:bold;font-size:1.1em}
.total-naranja{background:#e67e22!important;color:#fff!important;font-weight:bold}
.separador{border:none;border-top:1px dashed #ccc;margin:15px 0}
.pie{text-align:center;font-size:.8em;color:#999;margin-top:20px}
.tabla-premios-pdf{width:100%;border-collapse:collapse;margin-top:8px;font-size:.75em;color:#fff}
.tabla-premios-pdf th{border-bottom:1px solid rgba(255,255,255,.3);padding:3px 5px;text-align:left}
.tabla-premios-pdf td{border-bottom:1px solid rgba(255,255,255,.1);padding:2px 5px}
.mensajes-aliento{background:#f0f8ff;border-left:4px solid #1a3a5c;padding:15px;margin:15px 0;border-radius:8px}
.mensajes-aliento p{margin:5px 0;font-size:.9em}
</style>"""
    
    html = f"""<html><head><meta charset="UTF-8">{css}</head><body>
<h1>🎰 Quini 6 Checker</h1>
<p style="text-align:center">📌 {info['texto_completo']}</p>
<h2>🏆 NÚMEROS GANADORES</h2>
<div class="contenedor-modalidades">"""
    
    for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale"]:
        if mod in resultados:
            nums = resultados[mod]; pd = pozos.get(mod, {})
            monto = pd.get("monto","N/D") if isinstance(pd,dict) else str(pd)
            estado = pd.get("estado","?") if isinstance(pd,dict) else "?"
            gan = pd.get("ganadores",0) if isinstance(pd,dict) else 0
            ac = pd.get("aciertos_ganadores",0) if isinstance(pd,dict) else 0
            badge = '<span class="badge-vacante">⚠️ VACANTE</span>' if estado=="VACANTE" else (f'<span class="badge-ganado">✅ {gan} gan. ({ac} ac.)</span>' if (gan>0 and ac>0 and ac<6) else (f'<span class="badge-ganado">✅ {gan} gan.</span>' if gan>0 else '<span class="badge-ganado">✅ GANADO</span>'))
            bolitas = " ".join([f'<span class="bolita">{str(n).zfill(2)}</span>' for n in nums])
            tp_html = ""
            if isinstance(pd,dict) and "tabla_premios" in pd and pd["tabla_premios"]:
                tp_html = '<table class="tabla-premios-pdf"><tr><th>Premio</th><th>POZO $</th><th>Gan.</th><th>Premio $</th></tr>'
                for p in pd["tabla_premios"]:
                    tp_html += f"<tr><td>{p['premio']}</td><td>{p['pozo']}</td><td>{p['ganadores']}</td><td>{p['premio_ganador']}</td></tr>"
                tp_html += '</table>'
            html += f'<div class="tarjeta-modalidad"><h4>{mod}</h4><div>{bolitas}</div><div class="pozo-monto">🏆 {monto}</div><div>{badge}</div>{tp_html}</div>'
    
    html += "</div>"
    if "Premio Extra" in resultados:
        nums = resultados["Premio Extra"]; pd = pozos.get("Premio Extra", {})
        monto = pd.get("monto","N/D") if isinstance(pd,dict) else str(pd)
        gan = pd.get("ganadores",0) if isinstance(pd,dict) else 0
        badge = f'<span class="badge-ganado">✅ {gan} ganadores</span>' if gan>0 else '<span class="badge-ganado">✅ Con ganadores</span>'
        bolitas = " ".join([f'<span class="bolita-extra">{str(n).zfill(2)}</span>' for n in nums])
        html += f'<div class="tarjeta-extra"><h4>🎟️ Premio Extra</h4><div>{bolitas}</div><div class="pozo-monto">🏆 {monto}</div><div class="info-ganadores">{gan} ganadores</div><div>{badge}</div></div>'
    
    html += '<h2>💬 MENSAJES PARA LOS JUGADORES</h2><div class="mensajes-aliento">'
    for msg in mensajes_aliento:
        html += f'<p>{msg}</p>'
    html += '</div>'
    
    html += '<h2>📋 TABLERO DE JUGADAS</h2>'
    for jd in detalle:
        nombre = jd["nombre"]; nums_j = list(jd["modalidades"].values())[0]["numeros_jugados"]
        html += f'<h3>🔹 {nombre}</h3><table class="tabla-pdf"><thead><tr><th>Modalidad</th>'
        for n in nums_j: html += f'<th>{str(n).zfill(2)}</th>'
        html += '<th>Aciertos</th></tr></thead><tbody>'
        for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]:
            if mod in jd["modalidades"]:
                d = jd["modalidades"][mod]; a = d["aciertos"]
                if mod in ["Premio Extra", "Revancha"]:
                    ct = "total-verde" if a>=6 else ("total-naranja" if a>=3 else "")
                else:
                    ct = "total-verde" if a>=4 else ("total-naranja" if a>=3 else "")
                nm = "Extra" if mod=="Premio Extra" else mod
                html += f'<tr><td class="mod" style="color:#1a3a5c;background:#eaf0f8;font-weight:bold;text-align:left">{nm}</td>'
                for n in nums_j: html += '<td class="ok-pdf">OK</td>' if n in d["numeros_sorteados"] else '<td>-</td>'
                html += f'<td class="{ct}">{a}</td></tr>'
        html += '</tbody></table><br>'
    
    html += f'<hr class="separador"><p class="pie">Quini 6 Checker · {info["texto_completo"]} · Generado el {datetime.now().strftime("%d/%m/%Y %H:%M")}</p></body></html>'
    return html

# ---------- BARRA LATERAL ----------
with st.sidebar:
    st.header("⚙️ Configuración")
    tabs = st.tabs(["📋 Jugadas", "📧 Email", "🔍 Sorteos", "📥 Importar", "ℹ️ Info"])
    
    with tabs[0]:
        st.subheader("Tus jugadas")
        jugadas = cargar_jugadas()
        for i, j in enumerate(jugadas):
            with st.expander(f"🔹 {j['nombre']}"):
                nn = st.text_input("Nombre", value=j['nombre'], key=f"nom_{i}")
                email = st.text_input("Email", value=j.get("email", ""), key=f"email_{i}", placeholder="ejemplo@email.com")
                ns = ", ".join(str(n) for n in j["numeros"])
                nn2 = st.text_input("6 números (1-45)", value=ns, key=f"nums_{i}")
                try:
                    nl = [int(n.strip()) for n in nn2.split(",") if n.strip().isdigit()]
                    if len(nl)==6 and all(1<=n<=45 for n in nl): j["nombre"]=nn; j["numeros"]=nl; j["email"]=email
                    elif nn2!=ns: st.warning("6 números entre 1 y 45")
                except: st.error("Formato inválido")
                
                # Botón para eliminar esta jugada
                if st.button("🗑️ Eliminar jugada", key=f"del_{i}"):
                    jugadas.pop(i)
                    guardar_jugadas(jugadas)
                    st.success(f"✅ Jugada eliminada")
                    st.rerun()
        
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Guardar", use_container_width=True): guardar_jugadas(jugadas); st.success("¡Guardadas!")
        with c2:
            if st.button("➕ Nueva", use_container_width=True): jugadas.append({"nombre":f"Jugada {len(jugadas)+1}","email":"","numeros":[1,2,3,4,5,6]}); guardar_jugadas(jugadas); st.rerun()
    
    with tabs[1]:
        st.subheader("Configurar email")
        st.session_state["destinatario"] = st.text_input("Mail destinatario", value="adrian.bertalot@prysmian.com")
        st.session_state["remitente"] = st.text_input("Tu Gmail", value="bertaad736@gmail.com")
        st.session_state["password"] = st.text_input("Contraseña de app", type="password", value = "hmtw lcaq nlni ejqc")
        if st.button("📝 Guardar email"): st.success("Guardado")
    
    with tabs[2]:
        st.subheader("🔍 Sorteos anteriores")
        sg = obtener_sorteos_guardados()
        if sg:
            st.caption(f"📂 {len(sg)} sorteos")
            se = st.selectbox("Seleccioná:", ["Último sorteo"] + [f"Sorteo N° {s}" for s in sg])
            if st.button("📂 Cargar", use_container_width=True):
                if se == "Último sorteo":
                    with st.spinner("Cargando..."): res, poz, inf = obtener_resultados()
                else:
                    with st.spinner(f"Cargando N° {se.replace('Sorteo N° ','')}..."): res, poz, inf = obtener_resultados_por_numero(se.replace("Sorteo N° ",""))
                if res and len(res)>=4:
                    st.session_state.update({"resultados_cache":res, "pozos_cache":poz, "info_sorteo_cache":inf, "ultimo_chequeo":datetime.now().strftime("%d/%m %H:%M"), "mostrar_detalle":False})
                st.rerun()
        sb = st.text_input("N° de sorteo", placeholder="3368")
        if st.button("🔍 Buscar") and sb.strip().isdigit():
            with st.spinner(f"Buscando N° {sb}..."): res, poz, inf = obtener_resultados_por_numero(sb.strip())
            if res and len(res)>=4:
                st.session_state.update({"resultados_cache":res, "pozos_cache":poz, "info_sorteo_cache":inf, "ultimo_chequeo":datetime.now().strftime("%d/%m %H:%M"), "mostrar_detalle":False})
                st.rerun()
            else: st.error(f"No encontrado: {sb}")
    
    with tabs[3]:
        st.subheader("📥 Importar sorteos")
        af = st.file_uploader("Subir Excel", type=["xlsx"])
        if af:
            with open("temp.xlsx", "wb") as f:
                f.write(af.getbuffer())
            if st.button("📥 Importar"):
                cant, err = importar_desde_excel("temp.xlsx")
                if cant > 0:
                    st.success(f"✅ {cant} sorteos importados")
                    st.rerun()
                if err:
                    for e in err[:5]:
                        st.caption(f"• {e}")
        if st.button("📤 Exportar historial"):
            ok, r = exportar_historial_a_excel()
            if ok:
                st.success(r)
            else:
                st.error(r)
    
    with tabs[4]:
        st.info(f"📅 {obtener_fecha_sorteo_actual()}")
        st.caption("Sorteos: miércoles y domingos 21:15 hs")

# ============================================================
# PANEL PRINCIPAL
# ============================================================
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    if st.button("🎯 CARGAR ÚLTIMO SORTEO", type="primary", use_container_width=True):
        with st.spinner("🔍 Obteniendo..."): res, poz, inf = obtener_resultados()
        if res and len(res)>=4:
            st.session_state.update({"resultados_cache":res, "pozos_cache":poz, "info_sorteo_cache":inf, "ultimo_chequeo":datetime.now().strftime("%d/%m %H:%M"), "mostrar_detalle":False})
            st.rerun()
        else: st.error("No se pudieron obtener los resultados")

if st.session_state["resultados_cache"]:
    resultados = st.session_state["resultados_cache"]
    pozos = st.session_state["pozos_cache"]
    info = st.session_state["info_sorteo_cache"]
    
    st.markdown("---")
    st.markdown(f"## 🏆 NÚMEROS GANADORES — {info['texto_completo']}")
    
    mods_p = {k:v for k,v in resultados.items() if k!="Premio Extra"}
    mods_e = {k:v for k,v in resultados.items() if k=="Premio Extra"}
    cols = st.columns(4)
    
    for i, mod in enumerate(["Tradicional", "La Segunda", "Revancha", "Siempre Sale"]):
        if mod in mods_p:
            with cols[i]:
                nums = mods_p[mod]
                pd = pozos.get(mod,{})
                monto = pd.get("monto","N/D") if isinstance(pd,dict) else str(pd)
                estado = pd.get("estado","?") if isinstance(pd,dict) else "?"
                gan = pd.get("ganadores",0) if isinstance(pd,dict) else 0
                ac = pd.get("aciertos_ganadores",0) if isinstance(pd,dict) else 0
                
                if estado=="VACANTE": badge='<span class="badge-vacante">⚠️ VACANTE</span>'
                elif estado=="GANADO":
                    if gan>0:
                        if ac>0 and ac<6: badge=f'<span class="badge-ganado">✅ {gan} gan. ({ac} ac.)</span>'
                        else: badge=f'<span class="badge-ganado">✅ {gan} gan.</span>'
                    else: badge='<span class="badge-ganado">✅ GANADO</span>'
                else: badge='<span class="badge-desconocido">❓ Sin datos</span>'
                
                tp = ""
                if isinstance(pd,dict) and "tabla_premios" in pd and pd["tabla_premios"]:
                    tp = '<table style="width:100%;margin-top:10px;border-collapse:collapse;font-size:0.7em;color:#fff;"><tr style="border-bottom:1px solid rgba(255,255,255,0.3);"><th style="padding:3px 5px;text-align:left;">Premio</th><th style="padding:3px 5px;text-align:right;">POZO $</th><th style="padding:3px 5px;text-align:center;">Gan.</th><th style="padding:3px 5px;text-align:right;">Premio $</th></tr>'
                    for p in pd["tabla_premios"]:
                        tp += f'<tr style="border-bottom:1px solid rgba(255,255,255,0.1);"><td style="padding:2px 5px;text-align:left;">{p["premio"]}</td><td style="padding:2px 5px;text-align:right;">{p["pozo"]}</td><td style="padding:2px 5px;text-align:center;">{p["ganadores"]}</td><td style="padding:2px 5px;text-align:right;">{p["premio_ganador"]}</td></tr>'
                    tp += '</table>'
                
                bol = " ".join([f'<span class="numero-bola">{str(n).zfill(2)}</span>' for n in nums])
                st.markdown(f"""<div style="background:linear-gradient(135deg,#1a3a5c,#2c5aa0);padding:15px;border-radius:12px;text-align:center;color:#fff;margin-bottom:10px;box-shadow:0 4px 8px rgba(0,0,0,.3)"><div style="font-size:.9em;opacity:.9;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">{mod}</div><div style="margin:10px 0">{bol}</div><div style="font-size:.8em;margin-top:8px;opacity:.8">🏆 Pozo: {monto}</div><div style="margin-top:5px">{badge}</div>{tp}</div>""", unsafe_allow_html=True)
    
    if mods_e:
        st.markdown("---")
        for mod, nums in mods_e.items():
            pd = pozos.get(mod,{})
            monto = pd.get("monto","N/D") if isinstance(pd,dict) else str(pd)
            gan = pd.get("ganadores",0) if isinstance(pd,dict) else 0
            badge = f'<span class="badge-ganado">✅ {gan} ganadores</span>' if gan>0 else '<span class="badge-ganado">✅ Con ganadores</span>'
            bol = " ".join([f'<span class="numero-bola" style="font-size:.8em;width:35px;height:35px;line-height:35px">{str(n).zfill(2)}</span>' for n in nums])
            st.markdown(f"""<div style="background:linear-gradient(135deg,#6d1a8a,#9b59b6);padding:15px;border-radius:12px;text-align:center;color:#fff;margin-bottom:10px;box-shadow:0 4px 8px rgba(0,0,0,.3)"><div style="font-size:1em;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">🎟️ {mod}</div><div style="margin:10px 0">{bol}</div><div style="font-size:.8em;margin-top:8px;opacity:.8">🏆 Pozo: {monto}</div><div style="margin-top:5px">{badge}</div></div>""", unsafe_allow_html=True)
    
    st.markdown("---")
    col_d1, col_d2, col_d3 = st.columns([1, 2, 1])
    with col_d2:
        if not st.session_state["mostrar_detalle"]:
            if st.button("🔍 VER MIS JUGADAS", type="secondary", use_container_width=True): st.session_state["mostrar_detalle"]=True; st.rerun()
        else:
            if st.button("🔼 OCULTAR MIS JUGADAS", type="secondary", use_container_width=True): st.session_state["mostrar_detalle"]=False; st.rerun()
    
    if st.session_state["mostrar_detalle"]:
        st.markdown("---"); st.markdown("## 📋 TABLERO DE JUGADAS")
        jugadas = cargar_jugadas(); detalle = revisar_premios(jugadas, resultados)
        components.html(_construir_tablero_html(detalle, resultados), height=80+len(detalle)*310, scrolling=True)
        tp = sum(1 for j in detalle for mod, d in j["modalidades"].items() if (mod in ["Premio Extra", "Revancha"] and d["aciertos"]>=6) or (mod not in ["Premio Extra", "Revancha"] and d["aciertos"]>=4))
        if tp>0: st.balloons(); st.success(f"🚨 ¡TOTAL: {tp} premios!")
    
    st.markdown("---")
    st.markdown("### 📧 ENVIAR RESULTADOS POR MAIL")
    
    col_e1, col_e2, col_e3 = st.columns([1, 2, 1])
    with col_e2:
        # Clave de seguridad para envío de mails

        # Usar una variable de sesión para controlar el valor
        if "clave_valor" not in st.session_state:
            st.session_state["clave_valor"] = ""
        
        clave_envio = st.text_input(
            "🔑 Clave de seguridad", 
            type="password", 
            placeholder="Ingresá la clave para enviar", 
            value=st.session_state["clave_valor"],
            key="input_clave"
        )
        
        if st.button("✉️ ENVIAR POR MAIL A TODOS", use_container_width=True):
            if not st.session_state.get("password"): st.error("Configurá la contraseña en la barra lateral")
            elif clave_envio != "621512":
                st.error("❌ Clave incorrecta. No se enviaron los mails.")
            else:
                with st.spinner("📧 Preparando..."):
                    jugadas = cargar_jugadas()
                    detalle = revisar_premios(jugadas, resultados)
                    
                    mensajes_aliento = []
                    for jd in detalle:
                        nombre = jd["nombre"]
                        mejor_mod = max(jd["modalidades"].items(), key=lambda x: x[1]["aciertos"])
                        msg = _obtener_mensaje_aliento(
                            nombre,
                            sum(d["aciertos"] for d in jd["modalidades"].values()),
                            mejor_mod[0],
                            mejor_mod[1]["aciertos"],
                            info['numero']
                        )
                        mensajes_aliento.append(msg)
                    
                    html_completo = _construir_html_completo_con_mensajes(resultados, pozos, info, detalle, mensajes_aliento)
                    archivo_html = f"quini6_{info['numero'].replace('N° ','').strip()}.html"
                    with open(archivo_html, "w", encoding="utf-8") as f: f.write(html_completo)
                    
                    cuerpo_mail = f"🎯 Resultados Quini 6 - {info['texto_completo']}\n\n"
                    cuerpo_mail += f"Adjuntamos el archivo HTML con el detalle completo de todas las jugadas.\n\n"
                    cuerpo_mail += "━" * 40 + "\n"
                    cuerpo_mail += "💬 MENSAJES PARA LOS JUGADORES:\n"
                    cuerpo_mail += "━" * 40 + "\n\n"
                    for msg in mensajes_aliento:
                        cuerpo_mail += f"{msg}\n\n"
                    cuerpo_mail += "━" * 40 + "\n"
                    cuerpo_mail += "¡Saludos y mucha suerte para el próximo sorteo! 🍀\n"
                    cuerpo_mail += "━" * 40 + "\n"
                    
                    destinatarios = []
                    for j in jugadas:
                        email = j.get("email", "").strip()
                        if email: destinatarios.append(email)
                    if st.session_state.get("destinatario"): destinatarios.append(st.session_state["destinatario"])
                    destinatarios = list(set(destinatarios))
                    
                    if not destinatarios:
                        st.error("No hay emails configurados")
                    else:
                        destinatarios_str = ", ".join(destinatarios)
                        with st.spinner(f"📧 Enviando a {len(destinatarios)} destinatarios..."):
                            try:
                                enviar_correo_con_pdf(destinatarios_str, f"🎰 Quini 6 — {info['texto_completo']}", cuerpo_mail, archivo_html, st.session_state["remitente"], st.session_state["password"])
                                st.success(f"✅ Mail enviado a {len(destinatarios)} destinatarios")
                            except:
                                enviados = 0
                                for dest in destinatarios:
                                    try:
                                        enviar_correo_con_pdf(dest, f"🎰 Quini 6 — {info['texto_completo']}", cuerpo_mail, archivo_html, st.session_state["remitente"], st.session_state["password"])
                                        enviados += 1
                                    except Exception as e:
                                        st.error(f"❌ {dest}: {e}")
                                if enviados > 0:
                                    st.success(f"✅ {enviados} mails enviados")
                                    # Resetear la clave
                                    st.session_state["clave_valor"] = ""
                                    st.rerun()
else:
    st.info("👆 Hacé clic en **CARGAR ÚLTIMO SORTEO**")
    st.markdown("""<div style="text-align:center;padding:40px;background:linear-gradient(135deg,#1a3a5c,#2c5aa0);border-radius:20px;color:#fff;margin:20px 0"><div style="font-size:5em">🎰</div><div style="font-size:1.5em;margin:20px 0">Quini 6 Checker</div></div>""", unsafe_allow_html=True)

st.markdown("---")
st.caption(f"🕐 {st.session_state['ultimo_chequeo'] or 'Nunca'} | 📅 {obtener_fecha_sorteo_actual()}")