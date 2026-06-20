# Informe de Factibilidad - Database Nexus

**Proyecto:** Database Nexus: replicacion de datos entre bases de datos  
**Curso:** SI783 - Sistemas de Informacion  
**Lugar:** Tacna, Peru  
**Version:** 1.0  

## Descripcion del Proyecto

Database Nexus es una aplicacion web SaaS multiusuario para configurar, supervisar y ejecutar replicaciones por lotes entre PostgreSQL, MySQL, MariaDB, SQL Server, SQLite y MongoDB. El escenario de referencia es una pyme de Tacna que consolida ventas, inventario y clientes desde sistemas dispersos.

## Riesgos

| Riesgo | Impacto | Mitigacion |
|---|---|---|
| Credenciales externas mal gestionadas | Alto | Cifrado AES-256-GCM y variables de entorno |
| Cambios de esquema | Medio | Validacion previa y comparacion de esquemas |
| Volumen elevado | Medio | Lotes de 5000 filas y limites configurables |
| Datos personales | Alto | Cumplimiento de Ley 29733 |

## Factibilidad

El proyecto es tecnicamente viable por usar tecnologias maduras como React, Node.js, Fastify y PostgreSQL/Supabase. Es economicamente viable con una inversion estimada de S/ 12,980 y beneficio anual conservador de S/ 18,600. Operativamente requiere capacitacion breve para usuarios tecnicos y legalmente debe aplicar politicas de proteccion de datos personales.

## Analisis financiero

| Indicador | Valor |
|---|---:|
| Inversion inicial | S/ 12,980 |
| Beneficio anual estimado | S/ 18,600 |
| Relacion B/C | 1.43 |
| VAN referencial | S/ 3,120 |
| TIR referencial | 24% |

## Conclusion

Database Nexus es viable para un piloto peruano de integracion de datos. La solucion reduce reprocesos manuales, mejora trazabilidad y permite escalar hacia mecanismos de replicacion mas avanzados en versiones futuras.

