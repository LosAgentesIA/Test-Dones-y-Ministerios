#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera index.html con las afirmaciones del Excel embebidas.

Uso:
    python generar_datos.py

Solo usa la libreria estandar (zipfile + xml.etree). Reemplaza la marca
/*__DATOS__*/ de la plantilla por el JSON con las afirmaciones del archivo
Cuestionario Dones y Ministerios.xlsx.
"""
import json
import zipfile
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
XLSX = "Cuestionario Dones y Ministerios.xlsx"
SALIDA = "index.html"


def leer_afirmaciones(ruta):
    """Lee el xlsx y devuelve la lista de afirmaciones en orden de fila."""
    with zipfile.ZipFile(ruta) as z:
        sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
        compartidas = []
        for si in sst.findall("m:si", NS):
            textos = si.findall(".//m:t", NS)
            compartidas.append("".join(t.text or "" for t in textos))

        hoja = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        filas = hoja.findall(".//m:sheetData/m:row", NS)

    registros = []
    for fila in filas[1:]:  # saltar encabezado
        celdas = {}
        for c in fila.findall("m:c", NS):
            ref = c.get("r") or ""
            col = "".join(ch for ch in ref if ch.isalpha())
            tipo = c.get("t")
            v = c.find("m:v", NS)
            if v is None:
                continue
            valor = v.text
            if tipo == "s":
                valor = compartidas[int(valor)]
            celdas[col] = valor
        if "D" not in celdas:
            continue
        registros.append({
            "num": int(celdas.get("A") or 0),
            "categoria": (celdas.get("B") or "").strip(),
            "sub": (celdas.get("C") or "").strip(),
            "afirmacion": (celdas.get("D") or "").strip(),
        })
    return registros


PLANTILLA = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Test de Dones y Ministerios</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #f5f6f8;
    --superficie: #ffffff;
    --texto: #1b2433;
    --texto-suave: #5d6675;
    --borde: #e2e5ea;
    --acento: #3f6df0;
    --acento-fuerte: #2f57d0;
    --acento-suave: #e8eeff;
    --sombra: 0 12px 32px rgba(23, 34, 66, 0.08);
    --radio: 16px;
  }

  [data-tema="oscuro"] {
    --bg: #0f1522;
    --superficie: #1a2233;
    --texto: #eef1f7;
    --texto-suave: #a2abc0;
    --borde: #2a3448;
    --acento: #6f97ff;
    --acento-fuerte: #4f7bf0;
    --acento-suave: #232d45;
    --sombra: 0 12px 32px rgba(0, 0, 0, 0.4);
  }

  html { color-scheme: light; }
  [data-tema="oscuro"] { color-scheme: dark; }

  body {
    font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    background: var(--bg);
    color: var(--texto);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 24px 16px;
    transition: background 0.3s ease, color 0.3s ease;
  }

  header {
    width: 100%;
    max-width: 720px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 24px;
  }

  .marca {
    font-weight: 700;
    letter-spacing: 0.04em;
    color: var(--texto-suave);
    font-size: 0.82rem;
    text-transform: uppercase;
  }

  #btn-tema {
    background: var(--superficie);
    border: 1px solid var(--borde);
    border-radius: 999px;
    width: 44px;
    height: 44px;
    cursor: pointer;
    display: grid;
    place-items: center;
    color: var(--texto);
    transition: background 0.2s ease, transform 0.15s ease, border-color 0.2s ease;
  }
  #btn-tema:hover { border-color: var(--acento); transform: translateY(-1px); }
  #btn-tema svg { width: 20px; height: 20px; }
  .icono-sol { display: none; }
  .icono-luna { display: block; }
  [data-tema="oscuro"] .icono-sol { display: block; }
  [data-tema="oscuro"] .icono-luna { display: none; }

  main { width: 100%; max-width: 720px; }

  .vista { display: none; }
  .vista.activa {
    display: flex;
    flex-direction: column;
    align-items: stretch;
    animation: aparecer 0.35s ease both;
  }

  @keyframes aparecer {
    from { opacity: 0; transform: translateY(14px); }
    to { opacity: 1; transform: translateY(0); }
  }

  h1 { font-size: clamp(1.7rem, 5vw, 2.6rem); line-height: 1.15; text-align: center; }
  h2 { font-size: clamp(1.4rem, 4vw, 2rem); text-align: center; }
  h3 { font-size: 1.15rem; }

  .subtitulo { text-align: center; color: var(--texto-suave); font-size: 1.05rem; margin: 10px 0 28px; }

  .btn-primario {
    background: var(--acento);
    color: #ffffff;
    border: none;
    border-radius: 999px;
    padding: 16px 36px;
    font-size: 1.1rem;
    font-weight: 700;
    font-family: inherit;
    cursor: pointer;
    box-shadow: 0 8px 20px rgba(63, 109, 240, 0.28);
    transition: background 0.2s ease, transform 0.15s ease, box-shadow 0.2s ease;
  }
  .btn-primario:hover:not(:disabled) { background: var(--acento-fuerte); transform: translateY(-2px); }
  .btn-primario:active:not(:disabled) { transform: translateY(0); }
  .btn-primario:disabled { opacity: 0.45; cursor: not-allowed; box-shadow: none; }

  .btn-secundario {
    background: transparent;
    color: var(--texto);
    border: 1px solid var(--borde);
    border-radius: 999px;
    padding: 14px 28px;
    font-size: 1rem;
    font-weight: 600;
    font-family: inherit;
    cursor: pointer;
    transition: border-color 0.2s ease, background 0.2s ease, transform 0.15s ease;
  }
  .btn-secundario:hover:not(:disabled) { border-color: var(--acento); transform: translateY(-1px); }
  .btn-secundario:disabled { opacity: 0.45; cursor: not-allowed; }

  .acciones { display: flex; justify-content: space-between; gap: 12px; flex-wrap: wrap; }

  /* ---- Inicio ---- */
  #vista-inicio { align-items: center; justify-content: center; text-align: center; min-height: calc(100vh - 160px); gap: 8px; width: 100%; }
  #vista-inicio h1, #vista-inicio .subtitulo, #vista-inicio #btn-iniciar { margin-left: auto; margin-right: auto; }
  #vista-inicio .subtitulo { margin-bottom: 40px; max-width: 520px; }

  /* ---- Datos ---- */
  #vista-datos { gap: 14px; }
  #vista-datos h3 { margin-top: 12px; }

  .campo label { display: block; font-weight: 600; font-size: 1rem; margin-bottom: 8px; }
  .campo input {
    width: 100%;
    font-size: 1.15rem;
    padding: 14px 18px;
    border-radius: 12px;
    border: 2px solid var(--borde);
    background: var(--superficie);
    color: var(--texto);
    font-family: inherit;
    transition: border-color 0.2s ease;
  }
  .campo input:focus { outline: none; border-color: var(--acento); }

  .error { color: #e5484d; font-size: 0.92rem; margin-top: 6px; }
  .oculto { display: none !important; }

  .opciones-test { display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px; }

  .tarjeta-test {
    background: var(--superficie);
    border: 2px solid var(--borde);
    border-radius: var(--radio);
    padding: 20px 18px;
    text-align: center;
    cursor: pointer;
    font-size: 1rem;
    font-family: inherit;
    color: var(--texto);
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    min-height: 88px;
    transition: border-color 0.2s ease, transform 0.15s ease, box-shadow 0.2s ease, background 0.2s ease;
  }
  .tarjeta-test:hover { border-color: var(--acento); transform: translateY(-2px); box-shadow: var(--sombra); }
  .tarjeta-test.seleccionada { border-color: var(--acento); background: var(--acento-suave); box-shadow: var(--sombra); }
  .tarjeta-nombre { display: block; font-weight: 700; font-size: 1.1rem; text-align: center; margin-bottom: 0; }
  .tarjeta-info { color: var(--texto-suave); font-size: 0.9rem; }

  #vista-datos .btn-primario { align-self: center; margin-top: 18px; }

  /* ---- Test ---- */
  #vista-test { gap: 20px; }

  .progreso { display: flex; flex-direction: column; gap: 10px; }
  .progreso span { color: var(--texto-suave); font-size: 0.95rem; font-weight: 600; text-align: center; }

  .barra-fondo { background: var(--borde); border-radius: 999px; height: 10px; overflow: hidden; }
  .barra-llena { background: var(--acento); height: 100%; width: 0; border-radius: 999px; transition: width 0.4s ease; }

  .pregunta-card {
    background: var(--superficie);
    border: 1px solid var(--borde);
    border-radius: var(--radio);
    box-shadow: var(--sombra);
    padding: clamp(28px, 6vw, 48px);
    text-align: center;
    min-height: 200px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .pregunta-card.animar { animation: aparecer 0.35s ease both; }

  .afirmacion { font-size: clamp(1.25rem, 2.6vw, 1.7rem); line-height: 1.5; }

  .botones-respuesta { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 12px; }
  .btn-respuesta {
    background: var(--superficie);
    border: 2px solid var(--borde);
    border-radius: 14px;
    padding: 18px 12px;
    font-size: 1.1rem;
    font-weight: 700;
    font-family: inherit;
    color: var(--texto);
    cursor: pointer;
    transition: border-color 0.2s ease, background 0.2s ease, transform 0.15s ease;
  }
  .btn-respuesta:hover { border-color: var(--acento); transform: translateY(-2px); }
  .btn-respuesta.seleccionada { border-color: var(--acento); background: var(--acento-suave); }

  /* ---- Resultados ---- */
  #vista-resultados { gap: 20px; }

  .resultado-imprimible {
    background: var(--superficie);
    border: 1px solid var(--borde);
    border-radius: var(--radio);
    box-shadow: var(--sombra);
    padding: clamp(24px, 5vw, 40px);
  }
  .resultado-saludo { text-align: center; color: var(--texto-suave); margin: 8px 0 24px; }

  .selector-categorias { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; margin: 14px 0 22px; }
  .selector-categorias button {
    background: var(--superficie);
    border: 1px solid var(--borde);
    border-radius: 999px;
    padding: 10px 18px;
    font-size: 0.92rem;
    font-weight: 600;
    font-family: inherit;
    color: var(--texto);
    cursor: pointer;
    transition: all 0.2s ease;
  }
  .selector-categorias button:hover { border-color: var(--acento); transform: translateY(-1px); }
  .selector-categorias button.activo { background: var(--acento); color: #fff; border-color: var(--acento); box-shadow: 0 4px 12px rgba(63,109,240,0.25); }

  .grafico-bloque { margin-bottom: 34px; }
  .grafico-bloque:last-child { margin-bottom: 0; }
  .grafico-bloque h3 {
    margin-bottom: 14px;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--borde);
    color: var(--acento);
    text-align: left;
  }
  .grafico-contenedor { display: flex; flex-direction: column; align-items: center; gap: 14px; }
  .grafico-contenedor svg { width: 100%; max-width: 570px; height: auto; overflow: visible; }
  .leyenda-radar { display: flex; flex-wrap: wrap; gap: 8px 14px; justify-content: center; font-size: 0.86rem; color: var(--texto-suave); }
  .leyenda-item { display: inline-flex; align-items: center; gap: 6px; background: var(--acento-suave); border-radius: 999px; padding: 4px 10px; }
  .leyenda-punto { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
  .leyenda-item b { color: #0f172a; font-weight: 800; background: #fff; border-radius: 999px; padding: 2px 7px; margin-left: 2px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
  .leyenda-item.principal { background: var(--acento); color: #fff; }
  .leyenda-item.principal b { color: var(--acento); background: #fff; }
  .badge {
    background: var(--acento);
    color: #ffffff;
    border-radius: 999px;
    padding: 2px 10px;
    font-size: 0.7rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-left: 8px;
    vertical-align: middle;
    white-space: nowrap;
  }

  .pdf-membrete { display: none; }

  /* ---- Impresion ---- */
  @media print {
    body { background: #ffffff; padding: 0; color: #000000; }
    body * { visibility: hidden; }
    .resultado-imprimible, .resultado-imprimible * { visibility: visible; }
    .resultado-imprimible {
      position: absolute;
      left: 0;
      top: 0;
      width: 100%;
      border: none;
      box-shadow: none;
      padding: 0;
      background: transparent !important;
      color: #000000 !important;
    }
    .resultado-imprimible h2,
    .resultado-imprimible .resultado-saludo,
    .resultado-imprimible .grafico-bloque h3,
    .resultado-imprimible #resultado-nombre { color: #000000 !important; }
    .resultado-imprimible .grafico-bloque h3 { color: #3f6df0 !important; }
    .resultado-imprimible .leyenda-item { background: #eef2ff !important; color: #000 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .resultado-imprimible .leyenda-item.principal { background: #3f6df0 !important; color: #fff !important; }
    .resultado-imprimible .leyenda-item b { background: #fff !important; color: #0f172a !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .resultado-imprimible .leyenda-item.principal b { background: #fff !important; color: #3f6df0 !important; }
    .resultado-imprimible svg { -webkit-print-color-adjust: exact; print-color-adjust: exact; max-width: 440px !important; }
    .resultado-imprimible svg text { fill: #000000 !important; }
    .grafico-bloque.oculto { display: block !important; visibility: visible !important; }
    .pdf-membrete {
      display: block !important;
      visibility: visible !important;
      text-align: center;
      font-weight: 800;
      font-size: 1.05rem;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      color: #000000 !important;
      margin-bottom: 18px;
      padding-bottom: 12px;
      border-bottom: 2px solid #3f6df0;
    }
    .grafico-bloque { break-inside: avoid; page-break-inside: avoid; }
    .grafico-bloque:not(:last-child) { break-after: page; page-break-after: always; }
    .no-print { display: none !important; }
    header { display: none; }
  }
</style>
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/dist/umd/supabase.min.js"></script>
</head>
<body>

<header>
  <span class="marca">Instituto Ministerial del Espíritu Santo</span>
  <div style="display:flex; align-items:center; gap:12px;">
    <a href="admin.html" class="no-print" style="font-size:0.82rem; color:var(--texto-suave); text-decoration:none; border:1px solid var(--borde); padding:6px 12px; border-radius:999px;">Admin</a>
    <button id="btn-tema" class="no-print" aria-label="Cambiar modo claro u oscuro" title="Cambiar modo">
    <svg class="icono-sol" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true">
      <circle cx="12" cy="12" r="4"></circle>
      <path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"></path>
    </svg>
    <svg class="icono-luna" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"></path>
    </svg>
  </button>
  </div>
</header>

<main>

  <section id="vista-inicio" class="vista activa">
    <h1>Test de Dones y Ministerios</h1>
    <p class="subtitulo">Identifica los dones del Esp&iacute;ritu Santo y los ministerios en los que puedes servir mejor.</p>
    <button id="btn-iniciar" class="btn-primario">Iniciar Test</button>
  </section>

  <section id="vista-datos" class="vista">
    <h2>Datos del participante</h2>
    <div class="campo">
      <label for="input-nombre">Nombre</label>
      <input id="input-nombre" type="text" maxlength="40" placeholder="Tu nombre" autocomplete="given-name">
      <p id="error-nombre" class="error oculto">Por favor ingresa tu nombre.</p>
    </div>
    <div class="campo">
      <label for="input-apellido">Apellido</label>
      <input id="input-apellido" type="text" maxlength="40" placeholder="Tu apellido" autocomplete="family-name">
      <p id="error-apellido" class="error oculto">Por favor ingresa tu apellido.</p>
    </div>
    <div class="campo">
      <label for="input-email">Email</label>
      <input id="input-email" type="email" maxlength="100" placeholder="tu@email.com" autocomplete="email">
      <p id="error-email" class="error oculto">Por favor ingresa un email válido.</p>
    </div>
    <h3>Elige un test</h3>
    <div id="opciones-test" class="opciones-test"></div>
    <p id="error-test" class="error oculto">Por favor elige un test.</p>
    <button id="btn-comenzar" class="btn-primario">Comenzar</button>
  </section>

  <section id="vista-test" class="vista">
    <div class="progreso">
      <span id="pregunta-num">Pregunta 1 de 1</span>
      <div class="barra-fondo"><div id="progreso-barra" class="barra-llena"></div></div>
    </div>
    <div id="pregunta-card" class="pregunta-card">
      <p id="afirmacion" class="afirmacion"></p>
    </div>
    <div id="botones-respuesta" class="botones-respuesta"></div>
    <div class="acciones">
      <button id="btn-volver" class="btn-secundario">Volver</button>
      <button id="btn-siguiente" class="btn-primario" disabled>Siguiente</button>
    </div>
  </section>

  <section id="vista-resultados" class="vista">
    <div class="resultado-imprimible">
      <div class="pdf-membrete">Instituto Ministerial del Espíritu Santo</div>
      <h2>Resultados</h2>
      <p class="resultado-saludo">Gracias <strong id="resultado-nombre"></strong>, <span id="resultado-perfil">este es tu perfil de dones y ministerios:</span></p>
      <div id="selector-categorias" class="selector-categorias no-print oculto"></div>
      <div id="resultado-lista"></div>
    </div>
    <div class="acciones no-print">
      <button id="btn-pdf" class="btn-primario">Descargar PDF</button>
      <button id="btn-reiniciar" class="btn-secundario">Volver a empezar</button>
    </div>
  </section>

</main>

<script>
const DATOS = /*__DATOS__*/;

(function () {
  "use strict";

  const $ = (s) => document.querySelector(s);

  const ETIQUETAS = { "5 Ministerios de Jesucristo": "Los 5 Ministerios de Jes\u00fas", "Otros Dones y Ministerios": "Otros Dones y Ministerios", "Otros dones": "Otros Dones", "Otros ministerios": "Otros Ministerios", "Otros Dones": "Otros Dones", "Otros Ministerios": "Otros Ministerios" };
  const etiquetaCategoria = (cat) => ETIQUETAS[cat] || ("Los " + cat);

  function formatearLeyenda9Dones(sub) {
    const up = sub.toUpperCase();
    if (up.includes("PALABRA") && up.includes("SABIDURIA")) return "PALABRA DE SABIDUR\u00cdA";
    if (up.includes("PALABRA") && up.includes("REVELACION")) return "PALABRA DE REVELACI\u00d3N";
    return sub.replace(/^DON de /i, "").replace(/^DONES de /i, "").replace(/^DON del /i, "");
  }

  const ORDEN_FIJO = {};
  DATOS.forEach((d) => {
    let catKey = d.categoria;
    if (catKey === "Otros Dones y Ministerios") {
      catKey = d.sub.toUpperCase().trim().startsWith("DON") ? "Otros Dones" : "Otros Ministerios";
    }
    if (!ORDEN_FIJO[catKey]) ORDEN_FIJO[catKey] = [];
    if (!ORDEN_FIJO[catKey].includes(d.sub)) ORDEN_FIJO[catKey].push(d.sub);
  });

  const OPCIONES = [
    { valor: 3, etiqueta: "Mucho" },
    { valor: 2, etiqueta: "Algo" },
    { valor: 1, etiqueta: "Poco" },
    { valor: 0, etiqueta: "Nada" }
  ];

  let preguntas = [];
  let respuestas = [];
  let indice = 0;
  let nombre = "";
  let apellido = "";
  let email = "";
  let testSeleccionado = "";

  const VISTAS = ["inicio", "datos", "test", "resultados"];
  function mostrarVista(id) {
    VISTAS.forEach((v) => $("#vista-" + v).classList.toggle("activa", v === id));
    window.scrollTo(0, 0);
  }

  /* ---- Tema claro/oscuro ---- */
  function aplicarTema(t) {
    document.documentElement.dataset.tema = t;
    localStorage.setItem("tema", t);
  }
  document.getElementById("btn-tema").addEventListener("click", () => {
    aplicarTema(document.documentElement.dataset.tema === "oscuro" ? "claro" : "oscuro");
  });
  aplicarTema(localStorage.getItem("tema") || "claro");

  // ---- Supabase ----
  const SUPABASE_URL = "https://xstrpgjvflpajgcbzsai.supabase.co";
  const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhzdHJwZ2p2ZmxwYWpnY2J6c2FpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk0MDU5MjYsImV4cCI6MjEwNDk4MTkyNn0.hynO7EE8zSTznCiUJ-9VwCrtLOu50kw_qIKn4pLiG4M"; // Project Settings > API > anon public
  const supabaseClient = (window.supabase && SUPABASE_ANON_KEY) ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

  async function guardarResultado() {
    const email_norm = email.trim().toLowerCase();
    const nombreCompleto = (nombre + " " + apellido).trim();
    const puntajes = {};
    Object.keys(cachePorCategoria || {}).forEach((cat) => {
      puntajes[cat] = cachePorCategoria[cat].map((g) => ({ sub: g.sub, total: g.total, max: g.max }));
    });
    const payload = {
      nombre: nombreCompleto,
      email: email.trim(),
      email_norm,
      test_tipo: testSeleccionado,
      respuestas: [...respuestas],
      puntajes
    };
    if (!supabaseClient) {
      const hist = JSON.parse(localStorage.getItem("historial") || "[]");
      hist.push({ ...payload, fecha: new Date().toISOString() });
      localStorage.setItem("historial", JSON.stringify(hist));
      return;
    }
    try {
      await supabaseClient.from("personas").upsert({ nombre: nombreCompleto, email_norm }, { onConflict: "email_norm" });
      const { data: persona } = await supabaseClient.from("personas").select("id").eq("email_norm", email_norm).single();
      const persona_id = persona ? persona.id : null;
      await supabaseClient.from("resultados").insert({ persona_id, ...payload });
    } catch (e) {
      console.error("Supabase error, fallback localStorage", e);
      const hist = JSON.parse(localStorage.getItem("historial") || "[]");
      hist.push({ ...payload, fecha: new Date().toISOString() });
      localStorage.setItem("historial", JSON.stringify(hist));
    }
  }

  /* ---- Inicio ---- */
  document.getElementById("btn-iniciar").addEventListener("click", () => {
    construirOpciones();
    mostrarVista("datos");
  });

  /* ---- Datos ---- */
  function construirOpciones() {
    const cont = document.getElementById("opciones-test");
    cont.innerHTML = "";
    const tests = [
      { id: "9 Dones del Esp\u00edritu Santo", label: "Los 9 Dones del Esp\u00edritu Santo" },
      { id: "5 Ministerios de Jesucristo", label: "Los 5 Ministerios de Jes\u00fas" },
      { id: "Otros Dones", label: "Otros Dones" },
      { id: "Otros Ministerios", label: "Otros Ministerios" },
      { id: "COMPLETO", label: "Test Completo" }
    ];
    tests.forEach((t) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tarjeta-test";
      b.dataset.test = t.id;
      b.innerHTML = '<span class="tarjeta-nombre">' + t.label + '</span>';
      b.addEventListener("click", () => seleccionarTarjeta(b));
      cont.appendChild(b);
    });
  }

  function seleccionarTarjeta(tarjeta) {
    document.querySelectorAll(".tarjeta-test").forEach((t) => t.classList.toggle("seleccionada", t === tarjeta));
    testSeleccionado = tarjeta.dataset.test;
    document.getElementById("error-test").classList.add("oculto");
  }

  function seededRandom(seed) {
    let s = seed % 2147483647;
    if (s <= 0) s += 2147483646;
    return function() {
      s = (s * 16807) % 2147483647;
      return (s - 1) / 2147483646;
    };
  }

  function mezclarCompleto(datos) {
    const porCat = {
      "Otros Dones": [],
      "9 Dones del Esp\u00edritu Santo": [],
      "Otros Ministerios": [],
      "5 Ministerios de Jesucristo": []
    };
    datos.forEach((d) => {
      let catKey = d.categoria;
      if (d.categoria === "Otros Dones y Ministerios") {
        catKey = d.sub.toUpperCase().trim().startsWith("DON") ? "Otros Dones" : "Otros Ministerios";
      }
      if (!porCat[catKey]) porCat[catKey] = [];
      porCat[catKey].push(d);
    });
    const rng = seededRandom(20250824);
    Object.keys(porCat).forEach((cat) => {
      const arr = porCat[cat];
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(rng() * (i + 1));
        const tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
      }
    });
    const resultado = [];
    let ultimoCat = null;
    while (Object.values(porCat).some((a) => a.length > 0)) {
      let candidatos = Object.keys(porCat).filter((c) => porCat[c].length > 0 && c !== ultimoCat);
      if (candidatos.length === 0) {
        candidatos = Object.keys(porCat).filter((c) => porCat[c].length > 0);
      }
      candidatos.sort((a, b) => porCat[b].length - porCat[a].length || a.localeCompare(b));
      const elegido = candidatos[0];
      const item = porCat[elegido].shift();
      resultado.push(item);
      ultimoCat = elegido;
    }
    return resultado;
  }

  document.getElementById("btn-comenzar").addEventListener("click", () => {
    nombre = document.getElementById("input-nombre").value.trim();
    apellido = document.getElementById("input-apellido").value.trim();
    email = document.getElementById("input-email").value.trim();
    let ok = true;
    if (!nombre) {
      document.getElementById("error-nombre").classList.remove("oculto");
      ok = false;
    } else {
      document.getElementById("error-nombre").classList.add("oculto");
    }
    if (!apellido) {
      document.getElementById("error-apellido").classList.remove("oculto");
      ok = false;
    } else {
      document.getElementById("error-apellido").classList.add("oculto");
    }
    const emailValido = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    if (!email || !emailValido) {
      document.getElementById("error-email").classList.remove("oculto");
      ok = false;
    } else {
      document.getElementById("error-email").classList.add("oculto");
    }
    if (!testSeleccionado) {
      document.getElementById("error-test").classList.remove("oculto");
      ok = false;
    } else {
      document.getElementById("error-test").classList.add("oculto");
    }
    if (!ok) return;

    if (testSeleccionado === "COMPLETO") {
      preguntas = mezclarCompleto(DATOS);
    } else if (testSeleccionado === "Otros Dones") {
      preguntas = DATOS.filter((d) => d.categoria === "Otros Dones y Ministerios" && d.sub.toUpperCase().trim().startsWith("DON"));
    } else if (testSeleccionado === "Otros Ministerios") {
      preguntas = DATOS.filter((d) => d.categoria === "Otros Dones y Ministerios" && d.sub.toUpperCase().trim().startsWith("MINISTERIO"));
    } else {
      preguntas = DATOS.filter((d) => d.categoria === testSeleccionado);
    }
    respuestas = new Array(preguntas.length).fill(null);
    indice = 0;
    renderPregunta();
    mostrarVista("test");
  });

  /* ---- Preguntas ---- */
  function renderPregunta() {
    const p = preguntas[indice];
    document.getElementById("pregunta-num").textContent = "Pregunta " + (indice + 1) + " de " + preguntas.length;
    document.getElementById("progreso-barra").style.width = (((indice + 1) / preguntas.length) * 100) + "%";
    document.getElementById("afirmacion").textContent = p.afirmacion;

    const cont = document.getElementById("botones-respuesta");
    cont.innerHTML = "";
    OPCIONES.forEach((op) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "btn-respuesta" + (respuestas[indice] === op.valor ? " seleccionada" : "");
      b.textContent = op.etiqueta;
      b.dataset.valor = op.valor;
      b.addEventListener("click", () => seleccionarRespuesta(op.valor));
      cont.appendChild(b);
    });

    document.getElementById("btn-siguiente").disabled = respuestas[indice] === null;
    document.getElementById("btn-volver").disabled = indice === 0;

    const card = document.getElementById("pregunta-card");
    card.classList.remove("animar");
    void card.offsetWidth;
    card.classList.add("animar");
  }

  function seleccionarRespuesta(v) {
    respuestas[indice] = v;
    document.querySelectorAll(".btn-respuesta").forEach((b) => {
      b.classList.toggle("seleccionada", Number(b.dataset.valor) === v);
    });
    document.getElementById("btn-siguiente").disabled = false;
  }

  document.getElementById("btn-siguiente").addEventListener("click", () => {
    if (respuestas[indice] === null) return;
    if (indice < preguntas.length - 1) {
      indice++;
      renderPregunta();
    } else {
      mostrarResultados();
    }
  });

  document.getElementById("btn-volver").addEventListener("click", () => {
    if (indice > 0) {
      indice--;
      renderPregunta();
    }
  });

  /* ---- Resultados ---- */
  let cachePorCategoria = null;

  function acortarSub(sub) {
    const up = sub.toUpperCase();
    if (up.includes("PALABRA") && up.includes("SABIDURIA")) return "PALABRA DE SABIDUR\u00cdA";
    if (up.includes("PALABRA") && up.includes("REVELACION")) return "PALABRA DE REVELACI\u00d3N";
    return sub.replace(/^DON de /i,"").replace(/^DONES de /i,"").replace(/^MINISTERIO del /i,"").replace(/^MINISTERIO de /i,"").replace(/^DON del /i,"");
  }

  function renderRadar(cat, datos) {
    const n = datos.length; const cx = 180, cy = 180, R = 125;
    let svg = '<svg viewBox="0 0 360 360" class="svg-radar" role="img" aria-label="Radar ' + cat + '">';
    for (let g = 1; g <= 15; g++) {
      const r = R * g / 15;
      const isMajor = g % 5 === 0;
      svg += '<circle cx="' + cx + '" cy="' + cy + '" r="' + r + '" fill="none" stroke="var(--borde)" stroke-width="' + (isMajor ? 1 : 0.6) + '" opacity="' + (isMajor ? 0.55 : 0.22) + '"/>';
      if (isMajor) {
        svg += '<text x="' + (cx + 4) + '" y="' + (cy - r - 2) + '" font-size="7.5" font-weight="600" fill="#000000" opacity="1">' + g + "</text>";
      }
    }
    datos.forEach((d, i) => {
      const ang = (-90 + i * 360 / n) * Math.PI / 180;
      const x2 = cx + R * Math.cos(ang);
      const y2 = cy + R * Math.sin(ang);
      svg += '<line x1="' + cx + '" y1="' + cy + '" x2="' + x2 + '" y2="' + y2 + '" stroke="var(--borde)" stroke-width="1" opacity="0.3"/>';
    });
    const ceros = datos.filter((d) => d.total === 0).length;
    const usarFiltro = ceros >= 2;
    const activos = usarFiltro ? datos.filter((d) => d.total >= 1) : datos;
    if (activos.length >= 2) {
      const pts = activos.map((d) => {
        const i = datos.findIndex((x) => x.sub === d.sub);
        const ang = (-90 + i * 360 / n) * Math.PI / 180;
        const r = R * (d.total / d.max || 0);
        return (cx + r * Math.cos(ang)) + "," + (cy + r * Math.sin(ang));
      }).join(" ");
      svg += '<polygon points="' + pts + '" fill="rgba(63,109,240,0.18)" stroke="#3f6df0" stroke-width="2.2" stroke-linejoin="round" stroke-linecap="round"/>';
    } else if (!usarFiltro && datos.length >= 2) {
      const pts = datos.map((d, i) => {
        const ang = (-90 + i * 360 / n) * Math.PI / 180;
        const r = R * (d.total / d.max || 0);
        return (cx + r * Math.cos(ang)) + "," + (cy + r * Math.sin(ang));
      }).join(" ");
      svg += '<polygon points="' + pts + '" fill="rgba(63,109,240,0.18)" stroke="#3f6df0" stroke-width="2.2" stroke-linejoin="round" stroke-linecap="round"/>';
    }
    const puntosParaDibujar = usarFiltro ? activos : datos;
    puntosParaDibujar.forEach((d) => {
      const i = datos.findIndex((x) => x.sub === d.sub);
      const ang = (-90 + i * 360 / n) * Math.PI / 180;
      const r = R * (d.total / d.max || 0);
      const x = cx + r * Math.cos(ang);
      const y = cy + r * Math.sin(ang);
      const col = 'hsl(' + Math.round(i * 360 / n) + ', 72%, 52%)';
      svg += '<circle cx="' + x + '" cy="' + y + '" r="4.5" fill="' + col + '" stroke="#fff" stroke-width="1.6"/>';
    });
    datos.forEach((d, i) => {
      const ang = (-90 + i * 360 / n) * Math.PI / 180;
      const x = cx + (R + 18) * Math.cos(ang);
      const y = cy + (R + 18) * Math.sin(ang);
      const anchor = Math.cos(ang) > 0.35 ? "start" : Math.cos(ang) < -0.35 ? "end" : "middle";
      const label = acortarSub(d.sub);
      const fs = n > 12 ? "6.9" : "7.8";
      const needsWrap = label.length > 14 || label.includes(" Y ") || label.includes(" E ");
      if (needsWrap) {
        let parts;
        if (label.includes(" Y ")) { const p = label.split(" Y "); parts = [p[0], "Y " + p.slice(1).join(" Y ")]; }
        else if (label.includes(" E ")) { const p = label.split(" E "); parts = [p[0], "E " + p.slice(1).join(" E ")]; }
        else {
          const mid = Math.floor(label.length / 2);
          let sp = label.lastIndexOf(" ", mid);
          if (sp === -1) sp = label.indexOf(" ", mid);
          if (sp !== -1) parts = [label.slice(0, sp), label.slice(sp + 1)];
          else parts = [label];
        }
        if (parts.length === 2) {
          svg += '<text x="' + x + '" y="' + y + '" text-anchor="' + anchor + '" font-size="' + fs + '" font-weight="600" fill="var(--texto)">';
          svg += '<tspan x="' + x + '" dy="-0.45em">' + parts[0] + "</tspan>";
          svg += '<tspan x="' + x + '" dy="1.15em">' + parts[1] + "</tspan>";
          svg += "</text>";
        } else {
          svg += '<text x="' + x + '" y="' + y + '" text-anchor="' + anchor + '" dominant-baseline="middle" font-size="' + fs + '" font-weight="600" fill="var(--texto)">' + label + "</text>";
        }
      } else {
        svg += '<text x="' + x + '" y="' + y + '" text-anchor="' + anchor + '" dominant-baseline="middle" font-size="' + fs + '" font-weight="600" fill="var(--texto)">' + label + "</text>";
      }
    });
    svg += "</svg>";
    let html = '<div class="grafico-bloque" data-cat="' + cat + '"><h3>' + etiquetaCategoria(cat) + '</h3><div class="grafico-contenedor">' + svg + '<div class="leyenda-radar">';
    const ordenadosLeyenda = [...datos].sort((a, b) => b.total - a.total);
    ordenadosLeyenda.forEach((d, idxSorted) => {
      const principal = idxSorted === 0 ? " principal" : "";
      const fixedIdx = datos.findIndex((x) => x.sub === d.sub);
      const dotBg = 'hsl(' + Math.round(fixedIdx * 360 / datos.length) + ', 72%, 52%)';
      const textoSub = cat === "9 Dones del Esp\u00edritu Santo" ? formatearLeyenda9Dones(d.sub) : d.sub;
      html += '<span class="leyenda-item' + principal + '"><span class="leyenda-punto" style="background:' + dotBg + '"></span>' + textoSub + ' <b>' + d.total + "/" + d.max + "</b></span>";
    });
    html += "</div></div></div>";
    return html;
  }

  function mostrarResultados() {
    const porCategoria = {};
    preguntas.forEach((p, i) => {
      const v = respuestas[i] == null ? 0 : respuestas[i];
      let catKey = p.categoria;
      if (p.categoria === "Otros Dones y Ministerios") {
        catKey = p.sub.toUpperCase().trim().startsWith("DON") ? "Otros Dones" : "Otros Ministerios";
      }
      if (!porCategoria[catKey]) porCategoria[catKey] = [];
      let grupo = porCategoria[catKey].find((g) => g.sub === p.sub);
      if (!grupo) {
        grupo = { sub: p.sub, total: 0, max: 0 };
        porCategoria[catKey].push(grupo);
      }
      grupo.total += v;
      grupo.max += 3;
    });
    cachePorCategoria = porCategoria;
    document.getElementById("resultado-nombre").textContent = (nombre + " " + apellido).trim();
    const perfilMap = {
      "9 Dones del Espíritu Santo": "este es tu perfil de los 9 Dones del Espíritu Santo",
      "5 Ministerios de Jesucristo": "este es tu perfil de los 5 Ministerios de Jesucristo",
      "Otros Dones y Ministerios": "este es tu perfil de Otros Dones y Ministerios",
      "Otros Dones": "este es tu perfil de Otros Dones",
      "Otros Ministerios": "este es tu perfil de Otros Ministerios",
      "COMPLETO": "este es tu perfil de Dones y Ministerios"
    };
    document.getElementById("resultado-perfil").textContent = perfilMap[testSeleccionado] || "este es tu perfil de dones y ministerios:";
    guardarResultado();
    const cont = document.getElementById("resultado-lista");
    const selector = document.getElementById("selector-categorias");
    cont.innerHTML = "";
    selector.innerHTML = "";
    const esCompleto = testSeleccionado === "COMPLETO";
    let cats;
    if (esCompleto) {
      const ordenCompleto = ["5 Ministerios de Jesucristo", "9 Dones del Esp\u00edritu Santo", "Otros Dones", "Otros Ministerios"];
      cats = ordenCompleto.filter((c) => porCategoria[c]);
      if (cats.length === 0) cats = Object.keys(porCategoria);
    } else {
      if (porCategoria["Otros Dones"] && porCategoria["Otros Ministerios"]) {
        cats = ["Otros Dones", "Otros Ministerios"];
      } else {
        cats = Object.keys(porCategoria);
      }
    }
    if (esCompleto && cats.length > 1) {
      selector.classList.remove("oculto");
      cats.forEach((cat, idx) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.textContent = etiquetaCategoria(cat);
        btn.dataset.cat = cat;
        if (idx === 0) btn.classList.add("activo");
        selector.appendChild(btn);
      });
    } else {
      selector.classList.add("oculto");
    }
    cats.forEach((cat) => {
      const grupos = porCategoria[cat];
      const fijos = (ORDEN_FIJO[cat] || []).map((sub) => grupos.find((g) => g.sub === sub)).filter(Boolean);
      const datosParaRadar = fijos.length === grupos.length ? fijos : grupos;
      cont.insertAdjacentHTML("beforeend", renderRadar(cat, datosParaRadar));
    });
    if (esCompleto && cats.length > 1) {
      const bloques = cont.querySelectorAll(".grafico-bloque");
      bloques.forEach((b, i) => { if (i !== 0) b.classList.add("oculto"); });
      selector.querySelectorAll("button").forEach((btn) => {
        btn.addEventListener("click", () => {
          selector.querySelectorAll("button").forEach((b) => b.classList.toggle("activo", b === btn));
          const catSel = btn.dataset.cat;
          cont.querySelectorAll(".grafico-bloque").forEach((bloque) => {
            bloque.classList.toggle("oculto", bloque.dataset.cat !== catSel);
          });
        });
      });
    }
    mostrarVista("resultados");
  }

  document.getElementById("btn-pdf").addEventListener("click", () => window.print());

  document.getElementById("btn-reiniciar").addEventListener("click", () => {
    preguntas = [];
    respuestas = [];
    indice = 0;
    nombre = "";
    apellido = "";
    email = "";
    testSeleccionado = "";
    document.getElementById("input-nombre").value = "";
    const apellidoInput = document.getElementById("input-apellido");
    if (apellidoInput) apellidoInput.value = "";
    document.getElementById("input-email").value = "";
    document.querySelectorAll(".tarjeta-test").forEach((t) => t.classList.remove("seleccionada"));
    mostrarVista("inicio");
  });
})();
</script>
</body>
</html>
"""


def main():
    registros = leer_afirmaciones(XLSX)
    if not registros:
        raise SystemExit("No se encontraron afirmaciones en el Excel.")

    datos_json = json.dumps(registros, ensure_ascii=False, indent=1).replace("<", "\\u003c")
    html = PLANTILLA.replace("/*__DATOS__*/", datos_json)

    with open(SALIDA, "w", encoding="utf-8") as f:
        f.write(html)

    conteo = {}
    for r in registros:
        conteo[r["categoria"]] = conteo.get(r["categoria"], 0) + 1

    print("Generado:", SALIDA, "| total:", len(registros), "afirmaciones")
    for cat, n in conteo.items():
        print("  ", cat, "->", n)


if __name__ == "__main__":
    main()
