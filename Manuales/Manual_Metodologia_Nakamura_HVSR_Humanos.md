# Manual de Metodología de Análisis HVSR (Técnica Nakamura)
## Guía Integral para Ingenieros y Operadores Geofísicos
### Proyecto: PROMART – CITYMALL · ARDASYSMIC · 2026

---

## 1. Fundamentos Teóricos de la Técnica Nakamura (HVSR)

La técnica de la relación espectral entre las componentes horizontales y la componente vertical del ruido sísmico ambiental (HVSR, por sus siglas en inglés *Horizontal-to-Vertical Spectral Ratio*), propuesta originalmente por Yutaka Nakamura (1989), constituye uno de los métodos geofísicos no invasivos más eficaces y extendidos en la microzonificación sísmica y la ingeniería de sitios.

### 1.1 Naturaleza del Ruido Ambiental
El ruido sísmico ambiental (microtremores) está compuesto por una superposición de ondas de cuerpo y ondas superficiales generadas tanto por fuentes naturales de baja frecuencia ($f < 1\text{ Hz}$, como marejadas oceánicas y viento) como por fuentes antropogénicas de mayor frecuencia ($f > 1\text{ Hz}$, como tráfico rodado, maquinaria y actividad urbana).

En el modelo sismológico moderno:
- La componente vertical de los microtremores está fuertemente dominada por el movimiento elíptico retrógrado de las ondas Rayleigh. En la frecuencia fundamental de resonancia del estrato blando ($f_0$), el movimiento elíptico experimenta una inversión o singularidad, provocando un mínimo local relativo en el espectro vertical.
- Simultáneamente, las ondas Rayleigh y las ondas Love experimentan una amplificación máxima en el plano horizontal debido al confinamiento energético en la cubierta de baja velocidad.
- Al calcular el cociente espectral entre la media horizontal y la vertical:
  $$HVSR(f) = \frac{\sqrt{\frac{S_{NS}(f)^2 + S_{EW}(f)^2}{2}}}{S_V(f)}$$
  se obtiene una curva con un pico prominente a la frecuencia fundamental $f_0 = 1 / T_0$.

### 1.2 Relación de Cuarto de Longitud de Onda
Para un estrato homogéneo de espesor $H$ y velocidad de onda de corte promedio $V_s$ sobre un semiespacio rígido de roca:
$$T_0 = \frac{4H}{V_s} \quad \Longleftrightarrow \quad f_0 = \frac{V_s}{4H} \quad \Longleftrightarrow \quad H = \frac{V_s \cdot T_0}{4}$$

Esta relación permite estimar con notable precisión la profundidad del basamento competente cuando se conoce o se estima la velocidad de onda de corte promedio del depósito superficial.

### 1.3 Índice de Vulnerabilidad de Nakamura ($K_g$)
Nakamura (1997) formuló el índice $K_g$ como un estimador de la deformación por corte máxima elasto-plástica inducida en el estrato superficial ante eventos sísmicos severos:
$$K_g = \frac{A_0^2}{f_0} = A_0^2 \cdot T_0$$
Donde:
- $A_0$ es el factor de amplificación espectral relativo del pico fundamental.
- $f_0$ es la frecuencia fundamental en Hz.

Un valor de $K_g > 20$ se considera el umbral crítico a partir del cual el depósito de suelo es susceptible a experimentar no-linealidad dinámica, degradación severa de rigidez y deformaciones permanentes elevadas.

---

## 2. Protocolo de Operación y Adquisición en Campo

Para garantizar la reproducibilidad y confiabilidad de los datos, el operador de campo debe seguir estrictamente las siguientes directrices:

### 2.1 Equipos de Adquisición ARDASYSMIC
La red utiliza dos tipos de estaciones sísmicas digitales triaxiales:
1. **Estaciones Velocimétricas (`R13D7`):** Equipadas con geófonos triaxiales de velocidad (componentes `EHN`, `EHE`, `EHZ`). Son el transductor de elección preferencial para microtremores debido a su altísima sensibilidad en frecuencias bajas ($0.1\text{ Hz} \le f \le 2\text{ Hz}$).
2. **Estaciones Acelerométricas (`REAEC`):** Equipadas con acelerómetros triaxiales digitales MEMS (componentes `ENN`, `ENE`, `ENZ`). Son idóneas para eventos fuertes y vibraciones industriales, pero requieren tratamiento espectral especializado (`--units acceleration`) en microtremores para corregir la respuesta a bajas frecuencias.

> **Importante:** Por política de gobernanza técnica de ARDASYSMIC, está prohibido utilizar apodos informales o nombres de colores para referirse al hardware. El equipo siempre se designa por su código técnico (`R13D7` o `REAEC`).

### 2.2 Selección del Sitio y Acoplamiento
- Ubicar la estación en terreno natural firme y plano. Evitar carpetas asfálticas delgadas sueltas, gravilla movediza o rellenos heterogéneos no consolidados.
- Nivelar cuidadosamente la estación empleando el nivel de burbuja esférico incorporado.
- Orientar la marca del sensor hacia el Norte geográfico mediante brújula geológica (corrigiendo la declinación magnética local).
- Emplear protección contra el viento y la radiación solar directa (pantalla o campana protectora).
- Distanciar la estación al menos 20 a 30 metros de fuentes activas de ruido transitorio intenso (grupos electrógenos, zanjadoras, paso continuo de vehículos pesados).

### 2.3 Tiempo de Adquisición y Sincronización Temporal
- Duración mínima acumulada por sitio: **20 a 30 minutos** (mínimo 4 a 6 bloques continuos de 300 s).
- Llevar una bitácora exhaustiva registrando:
  - Hora de inicio y fin en **Reloj UTC Satelital**.
  - Hora interna mostrada por la interfaz del equipo DAS.
  - Coordenadas geográficas GNSS con al menos 6 decimales de precisión.
  - Condiciones meteorológicas y eventos antropogénicos identificables.

---

## 3. Procesamiento Computacional con TremorSys CLI

El procesamiento se ejecuta en la consola de comandos de Windows (PowerShell) mediante el binario oficial de **TremorSys CLI**.

### 3.1 Requisitos y Perfil de Procesamiento
Se debe contar con un archivo de perfil de configuración en formato TOML (`report_profile.toml`) que estipule:
- Rango de frecuencias: `0.20 - 20.0 Hz`.
- Longitud de ventanas: `25.0 s` con solapamiento de `50%`.
- Filtro de suavizado: `Konno-Ohmachi` con ancho de banda $b = 40$.
- Criterio de ventanas: Algoritmo adaptativo STA/LTA activado.

### 3.2 Paso 1: Procesamiento Individual
Cada archivo `.mseed` de 300 s se procesa por separado para evaluar la estabilidad de las señales:
```powershell
# Para archivos de estación velocimétrica R13D7:
tremorsys-cli --input "D:\Trabajos\Sitios Nakamura\N1\Registros\R13D7_182000_278.mseed" `
              --output "D:\Trabajos\.devtool\results_individual\N1_R13D7_182000" `
              --profile "D:\TremorSys\examples\report_profile.toml" `
              --no-record --no-noise --no-watermark

# Para archivos de estación acelerométrica REAEC:
tremorsys-cli --input "D:\Trabajos\Sitios Nakamura\N5\Registros\REAEC_132500_280.mseed" `
              --output "D:\Trabajos\.devtool\results_individual\N5_REAEC_132500" `
              --profile "D:\TremorSys\examples\report_profile.toml" `
              --units acceleration `
              --no-record --no-noise --no-watermark
```

### 3.3 Paso 2: Detección y Filtrado de Outliers
El operador inspecciona los archivos JSON resultantes (`__resultados.json`). Si un registro presenta:
- Período fundamental anómalo $T_0 \ge 2.0\text{ s}$ producido por tráfico pesado puntual o deriva térmica.
- Amplitudes excesivas $A_0 > 25$ aisladas no reproducidas por los demás registros del mismo sitio.
Dicho archivo debe ser catalogado como **candidato a descarte** y no debe incluirse en el agrupamiento principal.

### 3.4 Paso 3: Procesamiento Grupal Filtrado
Los archivos válidos que forman una secuencia temporal continua ($\Delta t \le 300\text{ s}$) se procesan conjuntamente para obtener la curva final de sitio:
```powershell
tremorsys-cli --group-input "D:\Trabajos\Sitios Nakamura\N1\Registros" `
              --group-size 5 `
              --output "D:\Trabajos\.devtool\results_group_filtered\N1" `
              --profile "D:\TremorSys\examples\report_profile.toml" `
              --no-record --no-noise --no-watermark
```

---

## 4. Control de Calidad SESAME (2004)

El estándar del proyecto europeo SESAME establece dos conjuntos de criterios obligatorios para validar una curva HVSR:

### 4.1 Criterios de Confiabilidad de la Curva (Reliability)
1. **Frecuencia Fundamental Significativa:** $f_0 > 10 / L_w$ (donde $L_w$ es la longitud de la ventana en segundos; para $L_w = 25\text{ s}$, $f_0 > 0.4\text{ Hz}$).
2. **Número de Ciclos Suficiente ($N_c$):** $N_c = l_w \cdot n_w \cdot f_0 > 200$, donde $n_w$ es el número de ventanas válidas seleccionadas.
3. **Baja Desviación Estándar:** La desviación estándar geométrica $\sigma_A(f)$ debe ser inferior a 2.0 en el rango $[0.5 f_0, 2 f_0]$.

### 4.2 Criterios de Claridad del Pico (Clarity)
1. Existe una frecuencia $f^-$ entre $f_0 / 4$ y $f_0$ tal que $A_0 / A(f^-) > 2$.
2. Existe una frecuencia $f^+$ entre $f_0$ y $4 f_0$ tal que $A_0 / A(f^+) > 2$.
3. La amplitud del pico fundamental cumple $A_0 > 2.0$.

### 4.3 Clasificación Simplificada SESAME
- **Clase A:** Pico destacado y confiable con $f_0 > 1.0\text{ Hz}$.
- **Clase B:** Pico destacado y confiable con $f_0 \le 1.0\text{ Hz}$ (típico de suelos blandos profundos).
- **Clase C:** Pico poco destacado ($1.5 \le A_0 \le 2.0$).
- **Clase D:** Curva plana o de baja amplitud ($1.0 \le A_0 < 1.5$, típico de roca o suelo muy rígido).
- **Clase E:** Amplitud $A_0 < 1.0$ (artefacto instrumental o inversión de campo de ondas).

---

## 5. Modelado 2D de Isoperíodos (Surfer / Matplotlib)

Una vez consolidados los períodos $T_0$ de los 15 sitios en la bitácora:
1. Exportar la terna espacial `(Longitud, Latitud, T0)` a un archivo de texto plano delimitado por comas (`DataSheet.dat`).
2. En Golden Software Surfer o en Python (Matplotlib + Scipy RBF/Cubic):
   - Generar una grilla regular de interpolación mediante método Kriging o Spline cúbico.
   - Definir contornos cada 0.05 s o 0.10 s en el rango de 0.8 s a 1.7 s.
   - Superponer la ubicación de los puntos N1 a N15 y las isolíneas de período.
   - Exportar en alta resolución (300 dpi) para el informe técnico (`Mapa_Isoperiodos_T0.png`).

---

## 6. Interpretación Geotécnica y Sismorresistente (Norma NEC-15)

Al redactar las conclusiones del informe, el profesional debe correlacionar los resultados geofísicos con la normativa vigente:

1. **Clasificación del Perfil de Suelo (NEC-SE-DS 2015):**
   - Suelos con $T_0 < 0.40\text{ s}$: Perfiles Tipo A, B o C (Roca o Suelo Denso/Rígido).
   - Suelos con $0.40\text{ s} \le T_0 < 0.75\text{ s}$: Perfil Tipo D (Suelo Medianamente Rígido).
   - Suelos con $T_0 \ge 0.75\text{ s}$ (y en particular $T_0 > 1.0\text{ s}$): **Perfil Tipo E (Suelo Muy Blando)**.
2. **Factores de Sitio y Espectro Elástico:**
   - Para Suelo Tipo E, los coeficientes de amplificación $F_a$ y $F_d$ aumentan notablemente.
   - La meseta espectral ($T_c$) y el período de transición ($T_L$) se desplazan hacia períodos largos, amplificando demandas sobre estructuras medianas y altas.
3. **Interacción Suelo-Estructura:**
   - Evaluar el período natural fundamental de la edificación ($T_n \approx 0.1 \cdot N$ para pórticos, o derivado de análisis modal espectral).
   - Si $|T_n - T_0| \le 0.2\text{ s}$, el riesgo de **resonancia destructiva** es crítico.
   - Recomendación: Modificar la rigidez estructural (incorporar arriostramientos o muros de corte) para alejar $T_n$ del período del suelo hacia la zona rígida ($T_n < 0.4\text{ s}$).
