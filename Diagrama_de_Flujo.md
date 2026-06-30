# Diagrama de Flujo — Smart Router (Sistema de Ruteo VS)

Este documento explica **cómo viajan los datos** dentro del sistema: desde que el
usuario sube el archivo Excel hasta que descarga las rutas optimizadas.

El sistema tiene **dos partes** que conversan entre sí:

| Parte | Archivo | Dónde corre | Rol |
|-------|---------|-------------|-----|
| **Front-end** (interfaz) | `Index.html` | En el **navegador** del usuario | Muestra el mapa, los controles y dibuja las rutas |
| **Back-end** (servidor) | `Code.gs` | En los **servidores de Google** | Lee el Excel, procesa datos y genera el archivo final |

La comunicación entre ambos se hace con `google.script.run`, que es el "puente"
que permite al navegador pedirle tareas pesadas al servidor.

---

## 1. Flujo general del sistema (visión completa)

```mermaid
flowchart TD
    A([Usuario abre la aplicación web]) --> B["doGet() entrega la página Index.html"]
    B --> C[Usuario sube el archivo Excel maestro]

    C -->|google.script.run| D["SERVIDOR: processUpload()<br/>Lee el Excel y cuenta pedidos<br/>por sucursal y horario"]
    D -->|Devuelve conteos| E[Se llenan los menús:<br/>Sucursal y Horarios]

    E --> F[Usuario elige sucursal, horarios<br/>y configura móviles, velocidad y tiempo]
    F --> G[Pulsa CALCULAR RUTAS]

    G -->|google.script.run| H["SERVIDOR: getOrders()<br/>Extrae los pedidos filtrados<br/>(cliente, dirección, coordenadas, bultos)"]
    H -->|Devuelve lista de pedidos| I["NAVEGADOR: procesar()<br/>Algoritmo de optimización por ángulo<br/>reparte pedidos entre rutas"]

    I --> J["render()<br/>Dibuja los puntos numerados en el mapa"]
    J --> K{¿El usuario quiere<br/>ajustar rutas?}

    K -->|Sí, manualmente| L[Reasigna pedidos:<br/>popup, casillas o lazo]
    L --> J
    K -->|No, está conforme| M[Pulsa EXPORTAR EXCEL]

    M -->|google.script.run| N["SERVIDOR: saveFinalRouting()<br/>Genera el Excel con colores,<br/>totales y auto-guardado en Drive"]
    N -->|Devuelve el archivo| O([Usuario descarga el<br/>Excel final con las rutas])

    style D fill:#be1877,color:#fff
    style H fill:#be1877,color:#fff
    style N fill:#be1877,color:#fff
    style I fill:#0c1b49,color:#fff
    style J fill:#0c1b49,color:#fff
    style O fill:#10b981,color:#fff
```

> 🟣 **Morado** = funciones del SERVIDOR (`Code.gs`)
> 🔵 **Azul** = funciones del NAVEGADOR (`Index.html`)
> 🟢 **Verde** = resultado final entregado al usuario

---

## 2. Detalle: ¿qué hace el servidor con el Excel? (`getOrders`)

Este es el paso clave donde los datos crudos del Excel se transforman en
pedidos listos para mapear.

```mermaid
flowchart TD
    A[Excel maestro en Google Drive] --> B[Copia temporal a Google Sheets<br/>para poder leer las celdas]
    B --> C[Limpia los encabezados con cleanTxt<br/>quita tildes y mayúsculas]
    C --> D[Ubica las columnas necesarias:<br/>LOCAL, ORDEN, DIRECCIÓN, LAT, LON...]

    D --> E{Por cada fila del Excel}
    E --> F{¿Coincide la sucursal<br/>y el horario elegido?}
    F -->|No| E
    F -->|Sí| G[Separa dirección y comuna por la coma]
    G --> H[Extrae el número de bultos con Regex<br/>busca 'Totales' + número]
    H --> I{¿Tiene FECHA DE PICKING?}
    I -->|Sí| J[sinArmar = 0<br/>pedido armado ✔]
    I -->|No| K[sinArmar = 1<br/>pedido pendiente ⚠]
    J --> L[Crea objeto del pedido]
    K --> L
    L --> M{¿Tiene coordenadas válidas?}
    M -->|No| N[Se descarta]
    M -->|Sí| O[Se agrega a la lista de pedidos]
    O --> E

    E -->|Fin de las filas| P[Devuelve la lista de pedidos<br/>al navegador]

    style A fill:#e0e0e0
    style P fill:#10b981,color:#fff
```

---

## 3. Detalle: el algoritmo de optimización (`procesar`)

El corazón del sistema. En vez de asignar rutas al azar, agrupa los pedidos
**geográficamente** para que cada móvil cubra una zona y los camiones no se crucen.

```mermaid
flowchart TD
    A[Lista de pedidos del servidor] --> B[Toma las coordenadas de la sucursal<br/>como punto de partida]
    B --> C["Por cada pedido calcula su ÁNGULO<br/>respecto a la sucursal<br/>(su posición en el 'reloj')"]
    C --> D[Ordena todos los pedidos por ángulo<br/>de modo que los vecinos queden juntos]
    D --> E["Divide la lista en N grupos iguales<br/>(N = cantidad de móviles)"]
    E --> F["Cada grupo = una RUTA<br/>(una 'tajada' geográfica del mapa)"]
    F --> G[Dentro de cada ruta, ordena<br/>del más cercano al más lejano]
    G --> H[Dibuja las rutas en el mapa]

    style A fill:#10b981,color:#fff
    style F fill:#d64394,color:#fff
    style H fill:#0c1b49,color:#fff
```

**Analogía sencilla para la presentación:** imagina la sucursal en el centro de
una pizza. El algoritmo corta la pizza en porciones (una por cada móvil), y cada
porción es una ruta. Así cada camión atiende su zona sin pisar la del otro.

---

## 4. Resumen de la "conversación" navegador ↔ servidor

```mermaid
sequenceDiagram
    participant U as 👤 Usuario
    participant N as 🌐 Navegador (Index.html)
    participant S as ⚙️ Servidor (Code.gs)
    participant D as 📁 Google Drive

    U->>N: Sube archivo Excel
    N->>S: processUpload(archivo)
    S->>D: Lee el Excel
    S-->>N: Conteo por sucursal/horario
    N-->>U: Muestra menús de selección

    U->>N: Elige sucursal + horarios + config
    N->>S: getOrders(sucursal, horarios)
    S->>D: Lee el Excel maestro
    S-->>N: Lista de pedidos
    N->>N: Optimiza rutas y dibuja el mapa
    N-->>U: Muestra rutas en el mapa

    U->>N: Ajusta rutas (opcional) y exporta
    N->>S: saveFinalRouting(rutas)
    S->>D: Guarda copia del Excel final
    S-->>N: Archivo Excel generado
    N-->>U: Descarga el Excel con las rutas
```

---

## Glosario rápido (por si lo preguntan en la defensa)

- **Front-end / Back-end:** la cara visible (lo que el usuario ve) vs. el motor
  oculto que hace los cálculos.
- **`google.script.run`:** el mecanismo que conecta el navegador con el servidor.
- **Regex (expresión regular):** una "fórmula de búsqueda" para encontrar texto
  con un patrón (aquí se usa para extraer el número de bultos).
- **Latitud / Longitud:** las coordenadas que ubican cada dirección en el mapa.
- **Fórmula de Haversine:** cálculo de distancia entre dos puntos del planeta
  considerando la curvatura de la Tierra.
- **Blob:** un bloque de datos binarios (la forma en que viaja un archivo).
- **Base64:** una forma de convertir un archivo a texto para poder enviarlo.
