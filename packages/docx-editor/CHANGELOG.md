# Changelog

Cambios de `@dimarborda/docx-editor`. El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/).

## [Unreleased]

## [0.12.1] - 2026-10-22

### Corregido

- `defineSection` y `sectionsFromHeadings` ya envuelven una sección existente que comparte el primer o el último párrafo con el rango nuevo; antes fallaban con «El rango corta otra sección…» ([#1](https://github.com/dimarborda/rust-web-docx/issues/1)).
