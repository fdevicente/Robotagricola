# Plan: Robot en servidor + Bot + Web privada con login

Fecha: 2026-07-29

## Objetivo
1. Sacar el robot del PC y dejarlo en un servidor 24/7 (lo más barato posible).
2. Mantener el bot de mensajería (Telegram hoy; evaluar WhatsApp).
3. Publicar una web con toda la info ordenada, con login y cumpliendo la ley
   chilena de protección de datos personales.

---

## 0. El bloqueador: el Excel MASTER

Todo lo demás depende de esto. Hoy el sistema gira en torno a un archivo
`MASTER…xlsx` que se abre y guarda entero con openpyxl. Eso no soporta:
- que el bot escriba mientras la web lee,
- que el usuario tenga el archivo abierto en Excel,
- backups consistentes ni acceso concurrente.

**Migrar a SQL (PostgreSQL) es requisito previo**, no un extra. Sin eso, el
servidor hereda los mismos problemas y agrega latencia.

**Tablas a modelar** (ya mapeadas): facturas, factura_lineas, cuenta_banco,
conciliaciones, proveedores, bitacora, inventario, vencimientos, personal,
vacaciones, tareas, cosechas, ajustes_manuales, boletas.

**Migración sin romper nada**: script que lee el Excel y puebla Postgres +
exportador a Excel para seguir entregando el archivo cuando se necesite.

---

## 1. Dónde alojarlo (comparativa)

| Opción | Costo aprox./mes | Pros | Contras |
|---|---|---|---|
| **Hetzner CX22** (2 vCPU, 4 GB, 40 GB) | **~€4,5 (~$5)** | El mejor precio/rendimiento; DPA disponible | Datacenter en Europa → ~200 ms a Chile |
| Vultr / DigitalOcean São Paulo | ~$6 | Cerca de Chile (~40 ms) | Un poco más caro |
| AWS Lightsail São Paulo | ~$5-7 | Cerca; ecosistema AWS | Más burocracia |
| Oracle Cloud Free Tier (ARM) | **$0** | Gratis y potente (4 vCPU/24 GB) | Pueden reclamar la instancia; sin SLA. No recomendado para producción |

**Recomendación: VPS en São Paulo (~$6/mes).** La diferencia con Hetzner es ~$1
y la latencia baja de 200 ms a 40 ms, que se nota en la web. Si el precio manda,
Hetzner es igual de válido (el bot no sufre latencia).

**Costo total estimado: ~US$7-9/mes** = VPS (~$6) + dominio (~$1) + backups (~$1).
Verificar precios al contratar; cambian.

---

## 2. Arquitectura

```
        Internet (HTTPS)
              │
        [ Caddy ]  ← TLS automático (Let's Encrypt)
              │
     ┌────────┴────────┐
     │   Web (Flask)   │  login + roles
     │   Bot Telegram  │  systemd, auto-restart
     └────────┬────────┘
              │
       [ PostgreSQL ]  ← disco cifrado
              │
     Backups cifrados diarios (off-site)
```

- **Sistema**: Ubuntu LTS. Servicios con **systemd** (sobrevive reinicios y
  crashes — reemplaza al watchdog de Windows).
- **HTTPS**: Caddy renueva certificados solo.
- **Firewall**: solo 80/443 y SSH con clave (sin contraseña, sin root).
- **Sin Ollama** (Claude extrae) → basta 2 GB de RAM.
- **El scraper del banco NO va al servidor**: las credenciales bancarias se
  quedan en el PC. La cartola se sube manualmente (ya funciona) o se corre el
  scraper en local y se sincroniza.

---

## 3. Mensajería: Telegram vs WhatsApp

| | Telegram | WhatsApp |
|---|---|---|
| Costo | **Gratis** | Por conversación (Meta Cloud API) |
| Requisitos | Ninguno | Cuenta Business verificada, número dedicado, aprobación de plantillas |
| Estado | **Ya funcionando** con Juan | Habría que construirlo |
| Archivos/fotos | Sí | Sí |

**Recomendación: mantener Telegram.** Ya funciona, es gratis y Juan lo usa.
Si más adelante otros trabajadores no quieren instalar Telegram, se agrega
WhatsApp vía **Meta Cloud API** reusando la misma lógica (el código de handlers
se abstrae de la mensajería). Ese sería un proyecto aparte.

---

## 4. La web privada

**Secciones** (reordenando lo que ya existe): Resumen · Conciliación ·
Facturas · Banco · Flujo de caja · Inventario y vencimientos · Bitácora y
jornadas · Maquinaria · Personal y vacaciones · Reportes.

**Login y roles** (mínimo necesario, no sobre-ingeniería):
- `dueño`: todo, incluidos sueldos y flujo.
- `capataz` (Juan): bitácora, inventario, tareas. **Sin sueldos ni banco.**
- `contador`: facturas, banco, reportes. Sin bitácora.

Implementación: Flask-Login + contraseñas con hash (bcrypt/argon2) + sesiones
seguras + **2FA (TOTP)** para el dueño + bloqueo tras intentos fallidos.

---

## 5. Protección de datos (Chile)

Se manejan **datos personales de trabajadores** (RUT, sueldos, vacaciones) y de
proveedores. Aplica la **Ley 19.628** y, sobre todo, la **Ley 21.719**, que crea
la Agencia de Protección de Datos y rige desde **diciembre de 2026** — es decir,
entra en vigencia dentro de este mismo año, así que conviene construir ya
cumpliendo. Verificar la fecha exacta y el detalle con un abogado.

**Medidas técnicas a implementar:**
- TLS en tránsito; disco y backups **cifrados** en reposo.
- Contraseñas hasheadas (nunca en texto plano); secretos en variables de
  entorno o gestor de secretos, **nunca en el repo**.
- **Principio de mínimo privilegio**: Juan no ve sueldos.
- **Registro de auditoría**: quién entró y qué consultó/modificó.
- Retención y borrado: política de cuánto se guarda y purga de lo que no se usa.
- Backups diarios cifrados fuera del servidor + **restauración probada**.
- Plan de respuesta ante brechas (la ley exige notificar).

**Medidas legales/organizativas:**
- Aviso de privacidad para los trabajadores (qué se recolecta y para qué).
- Registro de actividades de tratamiento.
- **DPA** con el proveedor de hosting (Hetzner/AWS/Vultr lo ofrecen).
- Si se aloja fuera de Chile, dejar documentada la transferencia internacional.

**Regla de oro**: las credenciales bancarias **no salen del PC local**.

---

## 6. Fases y orden

| Fase | Qué | Resultado |
|---|---|---|
| **A** | **Excel → PostgreSQL** (+ exportador a Excel) | Base sólida; requisito de todo lo demás |
| **B** | Login + roles + auditoría sobre la app actual | La web deja de ser "abierta en localhost" |
| **C** | VPS: Ubuntu, Docker/systemd, Caddy + HTTPS, dominio | Accesible desde internet, 24/7 |
| **D** | Bot Telegram al servidor (systemd) + backups cifrados | Se apaga el PC y todo sigue |
| **E** | Cumplimiento: aviso de privacidad, registro, DPA, prueba de restauración | Ordenado legalmente |
| **F** *(opcional)* | WhatsApp vía Meta Cloud API | Solo si se necesita |

**Se puede usar desde la Fase C**; D-E son para dejarlo robusto y en regla.

## Riesgos
- **Migración de datos**: correr Excel y SQL en paralelo unas semanas y comparar
  totales antes de apagar el Excel.
- **Costos**: revisar el gasto de la API de Claude al aumentar el uso.
- **Sin PC prendido**: el scraper del banco queda manual (ya resuelto con la
  carga de cartola).
