# 💜 Finanzas de pibas

App para ordenar tus finanzas: registrás **gastos**, **ingresos** e **inversiones**,
todo queda guardado en **tu propia planilla de Google**, y la app te muestra un
**análisis mensual** (cuánto entró, cuánto salió, en qué gastaste y cuánta plata
hay en cada cuenta).

👉 La app: **https://jpaltuna.github.io/finanzas-de-pibas/**

Cada usuaria tiene su planilla propia: nadie ve los datos de nadie.
La app corre en tu celu/PC y le habla directo a tu planilla.

---

## Cómo empezar a usarla (una sola vez, ~10 minutos)

### 1. Creá tu planilla
1. Entrá a [sheets.google.com](https://sheets.google.com) con tu cuenta de Google.
2. Creá una planilla nueva y ponele de nombre **Finanzas de pibas** (o el que quieras).

### 2. Pegá el conector
1. En la planilla, andá al menú **Extensiones → Apps Script**.
2. Borrá el código de ejemplo que aparece.
3. Abrí el archivo [`apps-script/conector.gs`](apps-script/conector.gs) de este proyecto,
   copiá **todo** el contenido y pegalo ahí.
4. Guardá (ícono del disquete 💾).

### 3. Implementalo como aplicación web
1. Botón azul **Implementar → Nueva implementación**.
2. En el engranaje ⚙️ elegí tipo **Aplicación web**.
3. Configurá:
   - **Ejecutar como:** Yo
   - **Quién tiene acceso:** Cualquier usuario
4. Tocá **Implementar**. Google te va a pedir autorización: es TU propio script
   escribiendo en TU propia planilla, dale permiso con tu cuenta.
   (Si aparece "Google no verificó esta app": **Configuración avanzada → Ir a … (no seguro)**.
   Es normal, pasa con todos los scripts caseros.)
5. Copiá la **URL de la aplicación web** (termina en `/exec`).

### 4. Conectá la app
1. Abrí **https://jpaltuna.github.io/finanzas-de-pibas/** en el celu o la PC.
2. Pegá la URL que copiaste y tocá **Probar conexión y guardar**.
3. ¡Listo! Las pestañas de la planilla (GASTOS, INGRESOS, INVERSIONES, CONFIG)
   se crean solas la primera vez.

### 5. Instalala como app
- **Android (Chrome):** menú ⋮ → **Agregar a la pantalla principal** (o "Instalar app").
- **iPhone (Safari):** botón compartir → **Agregar a inicio**.
- **PC (Chrome/Edge):** ícono de instalar 📥 que aparece en la barra de direcciones.

---

## Personalizar categorías y cuentas

Todo desde la app: en los desplegables, la opción **＋ Nueva…** agrega una
categoría o cuenta; y en la pantalla de configuración (⚙ arriba a la derecha)
podés **sacar con la ✕** las que no uses. También podés editar directamente la
pestaña **CONFIG** de tu planilla (la app las toma al tocar ↻).

## Gastos fijos y variables

Al cargar un gasto elegís si es **Fijo** (aparece todos los meses sí o sí:
alquiler, expensas, suscripciones, monotributo…) o **Variable** (donde vos
decidís cuánto). El análisis mensual te muestra cuánto se va en cada uno.

## ¿Y si cargo algo sin internet?

Queda guardado en el teléfono y se sube solo cuando vuelve la conexión
(o tocando "Reintentar ahora").

## Para actualizar el conector (si sale una versión nueva)

En Apps Script: pegá el código nuevo → guardar → **Implementar →
Administrar implementaciones** → lápiz ✏️ → Versión: **Nueva versión** →
**Implementar**. La URL `/exec` no cambia, no hay que tocar nada en la app.
