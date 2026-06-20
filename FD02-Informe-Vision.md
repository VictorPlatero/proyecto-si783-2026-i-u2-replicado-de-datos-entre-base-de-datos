# Documento de Vision - Database Nexus

**Sistema:** Database Nexus  
**Proyecto:** replicacion de datos entre bases de datos  
**Version:** 1.0  
**Lugar:** Tacna, Peru  

## Proposito

Definir la vision del producto Database Nexus, una plataforma web para replicar datos entre bases heterogeneas de forma segura, trazable y accesible para pymes peruanas.

## Alcance

Incluye autenticacion, configuraciones cifradas, replicacion por lotes, monitoreo de salud, documentacion de esquemas y administracion basica de usuarios. No incluye CDC exact-once en la primera version.

## Posicionamiento

El producto atiende a empresas que conviven con sistemas legados y bases diferentes. Frente a procesos manuales con CSV u hojas de calculo, Database Nexus permite configurar flujos repetibles y monitoreados.

## Usuarios

| Usuario | Necesidad |
|---|---|
| Gerencia | Reportes oportunos y confiables |
| Administrador TI | Control de conexiones y trabajos |
| Analista de datos | Consolidacion rapida de informacion |
| Usuario operativo | Datos actualizados para tareas diarias |

## Capacidades

- Autenticacion JWT.
- Credenciales cifradas.
- Conectores PostgreSQL, MySQL, MariaDB, SQL Server, SQLite y MongoDB.
- Modos insertar, upsert, reemplazar y recargar.
- Progreso persistente y reanudacion.
- Health monitor y documentador de esquemas.

## Conclusion

La vision es construir una herramienta realista, de costo moderado y alineada con necesidades peruanas de integracion de datos en organizaciones pequenas y medianas.

