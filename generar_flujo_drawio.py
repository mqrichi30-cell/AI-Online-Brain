# -*- coding: utf-8 -*-
import html

FONT='fontFamily=Inter, Helvetica, Arial;fontColor=#2D2E3A;'
S = {
 'start':'ellipse;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#10B981;strokeWidth=2;shadow=1;fontSize=11;fontStyle=1;spacing=16;'+FONT,
 'end':'ellipse;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#EF4444;strokeWidth=2;shadow=1;fontSize=11;fontStyle=1;spacing=16;'+FONT,
 'step':'rounded=1;arcSize=14;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#3B82F6;strokeWidth=2;shadow=1;fontSize=11;spacing=8;'+FONT,
 'dec':'rhombus;whiteSpace=wrap;html=1;fillColor=#FFFBEB;strokeColor=#F59E0B;strokeWidth=2;shadow=1;fontSize=10;spacing=30;spacingTop=4;spacingBottom=4;'+FONT,
 'sideend':'ellipse;whiteSpace=wrap;html=1;fillColor=#F8FAFC;strokeColor=#94A3B8;strokeWidth=2;shadow=1;fontSize=10;spacing=10;'+FONT+'fontColor=#64748B;',
 'sidestep':'rounded=1;arcSize=14;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#F97316;strokeWidth=2;shadow=1;fontSize=10;spacing=8;'+FONT,
}
# x, width per node type (container 700 wide; main column centered on x=220)
GX={'start':(75,290),'end':(75,290),'step':(65,310),'dec':(30,380),'sideend':(485,190),'sidestep':(460,230)}

CW=700; GAP=70; CX0=40; CY=110
PITCH=CW+GAP

cells=[]
def esc(t): return html.escape(t, quote=True)

def container(cid,x,y,w,h,title,fill,titlecolor):
    st=('rounded=1;arcSize=3;whiteSpace=wrap;html=1;fillColor=%s;strokeColor=none;'
        'verticalAlign=top;align=left;spacingLeft=18;spacingTop=12;fontSize=15;fontStyle=1;'
        'fontFamily=Inter, Helvetica, Arial;fontColor=%s;'%(fill,titlecolor))
    cells.append(f'<mxCell id="{cid}" value="{esc(title)}" style="{st}" vertex="1" parent="1"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/></mxCell>')

def edge(src,dst,label='',dashed=False,extra='',points=None):
    st=('edgeStyle=orthogonalEdgeStyle;rounded=1;html=1;jettySize=auto;'
        'strokeColor=#94A3B8;strokeWidth=2;endArrow=block;endFill=1;'
        'fontSize=10;fontStyle=1;fontFamily=Inter, Helvetica, Arial;fontColor=#475569;'
        'labelBackgroundColor=#FFFFFF;')
    if dashed:
        st=('edgeStyle=orthogonalEdgeStyle;rounded=1;html=1;jettySize=auto;'
            'dashed=1;dashPattern=6 4;strokeColor=#8B5CF6;strokeWidth=3;endArrow=block;endFill=1;'
            'fontSize=11;fontStyle=1;fontFamily=Inter, Helvetica, Arial;fontColor=#7C3AED;'
            'labelBackgroundColor=#FFFFFF;')
    st+=extra
    geo='<mxGeometry relative="1" as="geometry"/>'
    if points:
        pts=''.join(f'<mxPoint x="{px}" y="{py}"/>' for px,py in points)
        geo=f'<mxGeometry relative="1" as="geometry"><Array as="points">{pts}</Array></mxGeometry>'
    cells.append(f'<mxCell id="e_{src}_{dst}" value="{esc(label)}" style="{st}" edge="1" parent="1" source="{src}" target="{dst}">{geo}</mxCell>')

POS={}

def build_phase(idx,cid,fill,titlecolor,title,main,sides,edges_list):
    x=CX0+idx*PITCH
    y=75; layout={}
    for nid,typ,lab,h in main:
        layout[nid]=(typ,lab,y,h); POS[nid]=(idx,y,h)
        y+=h+60
    H=y+25
    container(cid,x,CY,CW,H,title,fill,titlecolor)
    for nid,(typ,lab,ny,h) in layout.items():
        gx,gw=GX[typ]
        cells.append(f'<mxCell id="{nid}" value="{esc(lab)}" style="{S[typ]}" vertex="1" parent="{cid}"><mxGeometry x="{gx}" y="{ny}" width="{gw}" height="{h}" as="geometry"/></mxCell>')
    for nid,typ,lab,h,ref in sides:
        rty,rlab,ry,rh=layout[ref]
        ny=ry+(rh-h)//2
        gx,gw=GX[typ]
        cells.append(f'<mxCell id="{nid}" value="{esc(lab)}" style="{S[typ]}" vertex="1" parent="{cid}"><mxGeometry x="{gx}" y="{ny}" width="{gw}" height="{h}" as="geometry"/></mxCell>')
    for e in edges_list: edge(*e)
    return H

cells.append('<mxCell id="banner" value="Automatización de correos AWG — Consolidaciones (Fases 00 → 06)" '
 'style="text;html=1;fontSize=24;fontStyle=1;fontFamily=Inter, Helvetica, Arial;fontColor=#1E293B;align=left;" '
 'vertex="1" parent="1"><mxGeometry x="48" y="26" width="1250" height="40" as="geometry"/></mxCell>')
cells.append('<mxCell id="legend" value="⬭ verde/rojo = inicio y fin · ▭ azul = paso · ▭ naranja = rama alterna · ◇ = decisión · ⬭ gris = correo ignorado · - - morado = pasa a la siguiente fase&#10;Carpeta de trabajo: Bandeja de entrada &gt; &quot;Marín, Cristhofer - AWG + Wakefern&quot; del buzón compartido pgcustservw2.im@pg.com · IA: ossgenai.im@pg.com · Seguimiento: lista SharePoint (sitio NACSO-RegionalVMI)" '
 'style="rounded=1;arcSize=10;whiteSpace=wrap;html=1;fillColor=#ffffff;strokeColor=#CBD5E1;shadow=1;align=left;spacingLeft=12;spacingRight=12;fontSize=11;fontFamily=Inter, Helvetica, Arial;fontColor=#475569;" '
 'vertex="1" parent="1"><mxGeometry x="1360" y="14" width="1500" height="64" as="geometry"/></mxCell>')

# ================= FASE 00 =================
p0_main=[
 ('p0_start','start','Entra un correo NUEVO a la carpeta "Marín, Cristhofer - AWG + Wakefern"',110),
 ('p0_trig','dec','¿Es correo nuevo de cliente?\n(no viene de la IA ni del sistema, sin marcas AWG y sin "No Match")',190),
 ('p0_ship','dec','¿El asunto es "AWG - Ship With POs Report"?',150),
 ('p0_ai','dec','¿Es correo de IA / sistema / Wakefern?',150),
 ('p0_recent','dec','¿Llegó en los últimos 60 minutos?',140),
 ('p0_mark','step','Marca el correo como LEÍDO y le pone categorías "RPA" y "Cristhofer Marín"',60),
 ('p0_dup','dec','¿Esta conversación ya fue ruteada antes? (registro en SharePoint)',170),
 ('p0_track','step','Crea registro de seguimiento en SharePoint con FlowId nuevo (Fase: 00-RouterSent)',60),
 ('p0_reply','step','Responde el correo original (mismo hilo) a la IA pidiendo clasificarlo: Consolidation / Wakefern Appointments / No Match',80),
 ('p0_move','step','Mueve el correo original a la subcarpeta "GENERAL - Email Route - 00"',60),
 ('p0_end','end','Espera la respuesta de la IA\n→ FASE 01',100),
]
p0_sides=[
 ('p0_ignore','sideend','Ignorado:\nno dispara el flujo',85,'p0_trig'),
 ('p0_block1','sideend','Bloqueado:\nno se procesa',85,'p0_ship'),
 ('p0_block2','sideend','Omitido: IA / sistema / ya ruteado',90,'p0_ai'),
 ('p0_old','sideend','Omitido:\ncorreo viejo',85,'p0_recent'),
 ('p0_dupend','sideend','Omitido:\nduplicado',85,'p0_dup'),
]
p0_edges=[
 ('p0_start','p0_trig'),
 ('p0_trig','p0_ship','Sí'),('p0_trig','p0_ignore','No'),
 ('p0_ship','p0_block1','Sí'),('p0_ship','p0_ai','No'),
 ('p0_ai','p0_block2','Sí'),('p0_ai','p0_recent','No'),
 ('p0_recent','p0_old','No'),('p0_recent','p0_mark','Sí'),
 ('p0_mark','p0_dup'),
 ('p0_dup','p0_dupend','Sí'),('p0_dup','p0_track','No'),
 ('p0_track','p0_reply'),('p0_reply','p0_move'),('p0_move','p0_end'),
]
build_phase(0,'c0','#F3EEFB','#6D28D9','FASE 00 · Router / Clasificador de correos',p0_main,p0_sides,p0_edges)

# ================= FASE 01 =================
p1_main=[
 ('p1_start','start','Llega un correo a la carpeta "Marín, Cristhofer - AWG + Wakefern"',110),
 ('p1_sender','dec','¿El remitente es la IA? (ossgenai.im@pg.com)',150),
 ('p1_parse','step','Lee del cuerpo el FlowId oculto (AWG-FLOW-ID) y el resultado: Consolidation / Wakefern Appointments / No Match',80),
 ('p1_find','dec','¿Hay registro con ese FlowId en Fase 00-RouterSent?',170),
 ('p1_update','step','Actualiza el registro con la clasificación (Fase: 00-RouterClassified)',60),
 ('p1_type','dec','¿La clasificación es "Consolidation"?',150),
 ('p1_consol','step','Marca el original "RPA" + "Follow up" (no leído) y lo mueve a "AWG Complete > AWG - CON - 01"; la respuesta de la IA también va a "AWG - CON - 01"',95),
 ('p1_asktable','step','Responde a la IA en el mismo hilo pidiendo la tabla Truck # / PO # (Fase: 01-IntakeSent)',75),
 ('p1_end','end','Espera la tabla de la IA\n→ FASE 02',100),
]
p1_sides=[
 ('p1_ignore','sideend','Ignorado:\nno es de la IA',85,'p1_sender'),
 ('p1_fallback','sidestep','Sin registro (fallback): busca el original en "GENERAL - Email Route - 00", raíz del cliente y "AWG Complete"; lo restaura y archiva la respuesta de la IA en "AWG - CON - 01"',150,'p1_find'),
 ('p1_return','sidestep','NO consolidación: devuelve el ORIGINAL a la carpeta raíz como NO leído y sin categorías; la respuesta de la IA se marca leída y va a "AWG - CON - 01". Fin.',140,'p1_type'),
]
p1_edges=[
 ('p1_start','p1_sender'),
 ('p1_sender','p1_ignore','No'),('p1_sender','p1_parse','Sí'),
 ('p1_parse','p1_find'),
 ('p1_find','p1_fallback','No'),('p1_find','p1_update','Sí'),
 ('p1_update','p1_type'),
 ('p1_type','p1_return','No'),('p1_type','p1_consol','Sí'),
 ('p1_consol','p1_asktable'),('p1_asktable','p1_end'),
]
build_phase(1,'c1','#EAF6EF','#047857','FASE 01 · Intake — lee la clasificación de la IA',p1_main,p1_sides,p1_edges)

# ================= FASE 02 =================
p2_main=[
 ('p2_start','start','Llega un correo de la IA (ossgenai.im@pg.com) a la carpeta del cliente',110),
 ('p2_track','dec','¿Hay registro con ese FlowId en Fase 01-IntakeSent?',170),
 ('p2_valid','dec','¿Trae una tabla Truck/PO válida?\n(no dice "Is not a consolidation" ni es tabla de datos)',190),
 ('p2_first','dec','¿Primera vez que se procesa esta consolidación?',160),
 ('p2_pos','dec','¿Se extrajeron 2 o más números de PO?',150),
 ('p2_ride','step','Envía el correo "Ride-With" en el mismo hilo del original, al buzón pgcustservw2.im@pg.com (Main PO + Riding-With POs)',85),
 ('p2_adv','step','Guarda los PO/SO en SharePoint (Fase: 02-ConsolTableRcvd) y mueve la respuesta de la IA a "AWG Complete"',75),
 ('p2_end','end','El correo "Ride-With" dispara la FASE 03',100),
]
p2_sides=[
 ('p2_cancel','sideend','Cancelado:\nno está en el flujo',85,'p2_track'),
 ('p2_restore','sidestep','No era consolidación: restaura el ORIGINAL a la carpeta raíz (no leído), archiva la respuesta de la IA en "AWG Complete" y BORRA el registro. Fin.',150,'p2_valid'),
 ('p2_dupend','sideend','Omitido:\nya procesada antes',90,'p2_first'),
 ('p2_novalid','sidestep','Menos de 2 PO: restaura el original a la carpeta raíz (no leído), archiva la respuesta en "AWG Complete" y BORRA el registro. Fin.',140,'p2_pos'),
]
p2_edges=[
 ('p2_start','p2_track'),
 ('p2_track','p2_cancel','No'),('p2_track','p2_valid','Sí'),
 ('p2_valid','p2_restore','No'),('p2_valid','p2_first','Sí'),
 ('p2_first','p2_dupend','No'),('p2_first','p2_pos','Sí'),
 ('p2_pos','p2_novalid','No'),('p2_pos','p2_ride','Sí'),
 ('p2_ride','p2_adv'),('p2_adv','p2_end'),
]
build_phase(2,'c2','#EBF2FC','#1D4ED8','FASE 02 · Recibe la tabla Truck/PO y pide Ride-With',p2_main,p2_sides,p2_edges)

# ================= FASE 03 =================
p3_main=[
 ('p3_start','start','Llega un correo al buzón compartido pgcustservw2.im@pg.com (se revisa cada 1 minuto)',120),
 ('p3_ride','dec','¿Es la notificación "Ride-With"?\n("scheduled to ride with another load")',180),
 ('p3_guard','dec','¿El FlowId aún NO fue procesado en esta fase?',160),
 ('p3_extract','step','Extrae el PO principal ("Main PO Number") y los PO que viajan con él ("Riding With POs")',70),
 ('p3_moc','step','Llena la plantilla MOC.xlsx con los PO (Office Script — SharePoint NACSO-RegionalVMI /General/Customer Documents/AWG-VMC/)',85),
 ('p3_send','step','Envía un correo nuevo "Awg Consolidation Request" con MOC.xlsx adjunto a nacsoshared.im@pg.com',75),
 ('p3_rdd','step','Responde en el hilo a la IA pidiendo alinear todos los PO a la última RDD con POM CODE L3 (AWG_RDD_ALIGNMENT)',80),
 ('p3_adv','step','Avanza el registro de SharePoint a Fase 03-RDDRequestSent',60),
 ('p3_arch','step','Mueve el correo Ride-With a "AWG Complete", lo marca leído y espera 1 hora',70),
 ('p3_end','end','Espera la respuesta de la IA\n→ FASE 04',100),
]
p3_sides=[
 ('p3_ignore','sideend','Ignorado:\nno aplica',85,'p3_ride'),
 ('p3_dup','sideend','Omitido:\nya procesado',85,'p3_guard'),
]
p3_edges=[
 ('p3_start','p3_ride'),
 ('p3_ride','p3_ignore','No'),('p3_ride','p3_guard','Sí'),
 ('p3_guard','p3_dup','No'),('p3_guard','p3_extract','Sí'),
 ('p3_extract','p3_moc'),('p3_moc','p3_send'),('p3_send','p3_rdd'),
 ('p3_rdd','p3_adv'),('p3_adv','p3_arch'),('p3_arch','p3_end'),
]
build_phase(3,'c3','#FDF1E7','#C2410C','FASE 03 · Ride-With — arma el MOC y pide alinear la RDD',p3_main,p3_sides,p3_edges)

# ================= FASE 04 =================
p4_main=[
 ('p4_start','start','Llega un correo al buzón compartido pgcustservw2.im@pg.com (se revisa cada 1 minuto)',120),
 ('p4_data','dec','¿Trae la tabla de datos de la IA?\n("Pom code", "L3", "Current RDD"…)',170),
 ('p4_rdd','step','Extrae la nueva RDD alineada y calcula la fecha de trabajo = RDD − 10 días',70),
 ('p4_find','dec','¿Hay registro de esta conversación en Fase 03-RDDRequestSent?',170),
 ('p4_window','dec','¿Hoy está dentro de la ventana de trabajo? (hoy ≥ RDD − 10 días)',180),
 ('p4_ready','step','Actualiza el registro: Status "Ready for OSS Follow Up" (Fase: 05-FullTableRequestSent) y pide a la IA la tabla completa de consolidación (PO, SO, Ship-to, RDD, Collective, Shipment, Delivery, Truck #)',115),
 ('p4_cat','step','Etiqueta el correo con "RPA", "Follow up", "Cristhofer Marín" y lo mueve a "AWG Complete"',75),
 ('p4_end','end','Espera la tabla completa de la IA\n→ FASE 05',100),
]
p4_sides=[
 ('p4_ignore','sideend','Ignorado:\nno aplica',85,'p4_data'),
 ('p4_cancel','sideend','Cancelado:\nno está en el flujo',90,'p4_find'),
 ('p4_pending','sidestep','Faltan más de 10 días: el registro queda en Status "Pending 10 Days Out" con la fecha de trabajo programada y avisa a la IA la fecha en que se trabajará',150,'p4_window'),
]
p4_edges=[
 ('p4_start','p4_data'),
 ('p4_data','p4_ignore','No'),('p4_data','p4_rdd','Sí'),
 ('p4_rdd','p4_find'),
 ('p4_find','p4_cancel','No'),('p4_find','p4_window','Sí'),
 ('p4_window','p4_ready','Sí'),('p4_window','p4_pending','No'),
 ('p4_ready','p4_cat'),('p4_pending','p4_cat'),('p4_cat','p4_end'),
]
build_phase(4,'c4','#FCF6E3','#A16207','FASE 04 · Respuesta de RDD y seguimiento',p4_main,p4_sides,p4_edges)

# ================= FASE 05 =================
p5_main=[
 ('p5_start','start','Llega un correo a la carpeta "Marín, Cristhofer - AWG + Wakefern"',110),
 ('p5_track','dec','¿Hay registro de esta conversación en Fase 05-FullTableRequestSent?',170),
 ('p5_table','dec','¿Trae la tabla completa de la IA?\n("Formatted Table" con PO #, SO #, RDD, Collective Number…)',190),
 ('p5_extract','step','Extrae la tabla HTML y, del encabezado del hilo, el remitente original y los demás destinatarios (para el To / CC)',85),
 ('p5_coll','dec','¿Algún "Collective Number" está en blanco o difiere entre las filas?',180),
 ('p5_notify','step','Envía al REMITENTE ORIGINAL (con CC a los demás) el correo "The following POs have been consolidated as requested" con la tabla',90),
 ('p5_phase','step','Avanza el registro de SharePoint a Fase 06-Consolidated',60),
 ('p5_arch','step','Etiqueta la respuesta de la IA ("RPA", "Follow Up", "AWG OSSGenAI Processed") y la mueve a "AWG Complete"',80),
 ('p5_end','end','FIN: cliente notificado de la consolidación',100),
]
p5_sides=[
 ('p5_cancel','sideend','Cancelado:\nno está en el flujo',90,'p5_track'),
 ('p5_ignore','sideend','Ignorado: no es la tabla completa',90,'p5_table'),
 ('p5_manual','sidestep','Todas las filas ya comparten Collective Number: avisa por correo a pgcustservw2.im@pg.com que ya está en la ventana de 10 días y debe revisarse y trabajarse MANUALMENTE',160,'p5_coll'),
]
p5_edges=[
 ('p5_start','p5_track'),
 ('p5_track','p5_cancel','No'),('p5_track','p5_table','Sí'),
 ('p5_table','p5_ignore','No'),('p5_table','p5_extract','Sí'),
 ('p5_extract','p5_coll'),
 ('p5_coll','p5_notify','Sí'),('p5_coll','p5_manual','No'),
 ('p5_notify','p5_phase'),('p5_phase','p5_arch'),
 ('p5_manual','p5_arch'),
 ('p5_arch','p5_end'),
]
build_phase(5,'c5','#FBEAEA','#B91C1C','FASE 05 · Notificación final de consolidación',p5_main,p5_sides,p5_edges)

# ================= FASE 06 =================
p6_main=[
 ('p6_start','start','TEMPORIZADOR: se ejecuta solo, cada 30 minutos (no depende de correos)',120),
 ('p6_cutoff','step','Calcula la hora de corte: ahora − 5 horas',60),
 ('p6_get','step','Busca en SharePoint registros atascados: fase distinta de 06-Consolidated y 99-ManualReview, sin avanzar desde antes del corte',95),
 ('p6_stuck','dec','¿Hay registros atascados?\n(más de 5 h sin avanzar)',160),
 ('p6_alert','step','Por cada uno: envía alerta a pgcustservw2.im@pg.com — "[AWG CON] Flujo de consolidación atascado (+5h)" con FlowId, fase, POs y asunto original',100),
 ('p6_mark','step','Marca cada registro atascado como Fase 99-ManualReview (revisión manual)',70),
 ('p6_clean','step','Busca registros ya terminados (Fase 06-Consolidated) y BORRA esas filas de la lista de seguimiento',85),
 ('p6_end','end','FIN del barrido (vuelve a ejecutarse en 30 minutos)',100),
]
p6_sides=[]
p6_edges=[
 ('p6_start','p6_cutoff'),('p6_cutoff','p6_get'),('p6_get','p6_stuck'),
 ('p6_stuck','p6_alert','Sí'),
 ('p6_stuck','p6_clean','No','','exitX=1;exitY=0.5;exitDx=0;exitDy=0;entryX=1;entryY=0.5;entryDx=0;entryDy=0;'),
 ('p6_alert','p6_mark'),('p6_mark','p6_clean'),('p6_clean','p6_end'),
]
# fix tuple shapes: edge(src,dst,label,dashed,extra)
p6_edges=[
 ('p6_start','p6_cutoff'),('p6_cutoff','p6_get'),('p6_get','p6_stuck'),
 ('p6_stuck','p6_alert','Sí'),
 ('p6_stuck','p6_clean','No',False,'exitX=1;exitY=0.5;exitDx=0;exitDy=0;entryX=1;entryY=0.5;entryDx=0;entryDy=0;'),
 ('p6_alert','p6_mark'),('p6_mark','p6_clean'),('p6_clean','p6_end'),
]
build_phase(6,'c6','#EDEFF3','#334155','FASE 06 · Barrido de flujos atascados (mantenimiento)',p6_main,p6_sides,p6_edges)

# ===== cross-phase connectors through corridors =====
def cross(src,dst,label):
    si,sy,sh=POS[src]; di,dy,dh=POS[dst]
    src_cy=CY+sy+sh//2
    dst_cy=CY+dy+dh//2
    midx=CX0+si*PITCH+CW+GAP//2
    edge(src,dst,label,dashed=True,
         extra='exitX=1;exitY=0.5;exitDx=0;exitDy=0;entryX=0;entryY=0.5;entryDx=0;entryDy=0;',
         points=[(midx,src_cy),(midx,dst_cy)])

cross('p0_end','p1_start','La IA responde la clasificación')
cross('p1_end','p2_start','La IA responde con la tabla Truck/PO')
cross('p2_end','p3_start','Llega el correo "Ride-With"')
cross('p3_end','p4_start','La IA responde con la RDD alineada')
cross('p4_end','p5_start','La IA responde con la tabla completa')

body='\n'.join(cells)
xml=f'''<mxfile host="app.diagrams.net" version="24.0.0">
  <diagram id="awg-email-flow" name="Flujo de correos AWG">
    <mxGraphModel dx="1600" dy="900" grid="0" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="0" pageScale="1" pageWidth="850" pageHeight="1100" math="0" shadow="0" background="#F8FAFC">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        {body}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>'''
open('/home/user/AI-Online-Brain/Flujo_Correos_AWG.drawio','w').write(xml)
print("OK cells:",len(cells))
