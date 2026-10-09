# 📊 Reto Integrador - Sesión 1 (Chuleta Rápida)

## Objetivo

Limpiar, validar y analizar una base de ventas utilizando Excel.

---

# 1. Limpiar nombres de vendedores

Eliminar espacios al inicio/final y convertir a formato correcto.

```excel
=NOMPROPIO(ESPACIOS(C2))
```

### Ejemplos

| Original | Resultado |
|-----------|-----------|
| ana maria | Ana Maria |
| PABLO | Pablo |
| MARTA LOPEZ | Marta Lopez |

---

# 2. Limpiar región

```excel
=NOMPROPIO(ESPACIOS(D2))
```

### Opcional: convertir MTY → Monterrey

```excel
=SI(NOMPROPIO(ESPACIOS(D2))="Mty";"Monterrey";NOMPROPIO(ESPACIOS(D2)))
```

---

# 3. Limpiar producto

```excel
=NOMPROPIO(ESPACIOS(E2))
```

---

# 4. Estandarizar estatus

Convertir todos los registros a mayúsculas.

```excel
=MAYUSC(ESPACIOS(H2))
```

### Resultado

| Original | Resultado |
|-----------|-----------|
| pendiente | PENDIENTE |
| completado | COMPLETADO |
| COMPLETADO | COMPLETADO |

---

# 5. Convertir MONTO a número

Eliminar:

- $
- MXN
- espacios
- separadores de miles

```excel
=VALOR(SUSTITUIR(SUSTITUIR(SUSTITUIR(ESPACIOS(G2);"$";"");"MXN";"");",";""))
```

### Ejemplos

| Original | Resultado |
|-----------|-----------:|
| $1,250.00 | 1250 |
| 980 MXN | 980 |
| $4,900.00 | 4900 |
| 1,500 MXN | 1500 |

---

# 6. Convertir fechas

## Caso simple (Excel ya reconoce la fecha)

```excel
=FECHA(AÑO(B2);MES(B2);DIA(B2))
```

Aplicar formato:

```text
dd/mm/yyyy
```

---

## Caso texto

```excel
=FECHANUMERO(B2)
```

---

## Si existen varios formatos mezclados

Ejemplos:

```text
12/03/2024
2024-04-09
18-03-2024
03/15/2025
```

Utilizar Power Query o convertir previamente todas las fechas a un mismo formato.

---

# 7. Detectar duplicados

```excel
=CONTAR.SI($A:$A;A2)
```

### Interpretación

- 1 = único
- >1 = duplicado

---

# 8. Total de ventas

```excel
=SUMA(G2:G26)
```

---

# 9. Promedio de ventas

```excel
=PROMEDIO(G2:G26)
```

---

# 10. Conteo de registros

```excel
=CONTARA(A2:A26)
```

---

# 11. Conteo de pendientes

```excel
=CONTAR.SI(H2:H26;"PENDIENTE")
```

---

# 12. Conteo de completados

```excel
=CONTAR.SI(H2:H26;"COMPLETADO")
```

---

# 13. Total de ventas por región

Supongamos que la región está en K2:

```excel
=SUMAR.SI($D:$D;K2;$G:$G)
```

Ejemplo:

| Región |
|---------|
| Guadalajara |
| Tijuana |
| Mexicali |
| Monterrey |

---

# 14. Región con mayores ventas

## Obtener el valor más alto

```excel
=MAX(L2:L5)
```

## Obtener el nombre de la región

```excel
=INDICE(K2:K5;COINCIDIR(MAX(L2:L5);L2:L5;0))
```

---

# 15. Formato condicional

## Ventas mayores a $5,000

Fórmula:

```excel
=G2>5000
```

Color recomendado:

🟢 Verde

---

## Estatus pendiente

Fórmula:

```excel
=H2="PENDIENTE"
```

Color recomendado:

🟡 Amarillo o 🔴 Rojo

---

# Estructura final esperada

| ID | FECHA | VENDEDOR | REGION | PRODUCTO | CANTIDAD | MONTO | ESTATUS |
|-----|---------|------------|----------|------------|----------|----------|-----------|
| VT-001 | 12/03/2024 | Ana Maria | Guadalajara | Laptop Lenovo | 2 | 1250 | PENDIENTE |

---

# Resumen rápido para la sesión

## Limpiar texto

```excel
=ESPACIOS()
=NOMPROPIO()
=MAYUSC()
```

## Limpiar montos

```excel
=SUSTITUIR()
=VALOR()
```

## Totales

```excel
=SUMA()
=PROMEDIO()
```

## Conteos

```excel
=CONTARA()
=CONTAR.SI()
```

## Ventas por región

```excel
=SUMAR.SI()
```

## Región ganadora

```excel
=MAX()
=INDICE()
=COINCIDIR()
```

## Duplicados

```excel
=CONTAR.SI()
```

✅ Limpieza de datos  
✅ Estandarización de nombres  
✅ Conversión de fechas  
✅ Conversión de montos  
✅ Eliminación de duplicados  
✅ Formato condicional  
✅ Totales y promedios  
✅ Ventas por región  
✅ Identificación de la mejor zona de ventas