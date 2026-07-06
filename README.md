# Cuadro de Control de Calidad — Recepción de Materiales

Aplicación web para generar automáticamente el cuadro de control de calidad de recepción de materiales en obras de edificación, basado en la **Plantilla RPC Rev. 81 (Febrero 2026)** del Colegio Oficial de Aparejadores, Arquitectos Técnicos e Ingenieros de Edificación de Madrid.

## Uso

1. Introduce el nombre de la obra
2. Adjunta el PDF de mediciones (presupuesto sin precios)
3. Descarga el Excel generado automáticamente

## Características

- Cruce automático entre partidas de la medición y la plantilla RPC
- Nomenclatura propia de la obra (no texto genérico de la plantilla)
- Sin duplicados: cada material aparece una sola vez por capítulo
- Normas fusionadas en una sola fila cuando un material tiene varias
- Capitalización normalizada según reglas gramaticales del español
- Dos hojas: cuadro de trabajo en obra + referencia completa RPC con leyenda de códigos

## Estructura del Excel generado

**Hoja 1 — Control Recepción Materiales:** Capítulo · Producto/Material · Marca · Modelo · Norma · CE · DdP · Otros controles · Partidas asociadas · (campos amarillos para rellenar en obra)

**Hoja 2 — Plantilla RPC Completa:** Lista íntegra de la plantilla RPC + leyenda de códigos de «Otros controles»

## Referencia normativa

Plantilla RPC Rev. 81 · Febrero 2026 · Colegio Oficial de Aparejadores, Arquitectos Técnicos e Ingenieros de Edificación de Madrid
